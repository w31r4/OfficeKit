import { createHash, createHmac, randomUUID, timingSafeEqual } from "node:crypto";
import { createServer } from "node:https";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { appendAuditRecord } from "./state.mjs";
import { excelLiveError, toExcelLiveFailure } from "./errors.mjs";
import {
  createExcelFailure,
  createExcelSuccess,
  EXCEL_LIVE_PROTOCOL,
  MAX_IMAGE_BYTES,
  MAX_REQUEST_BYTES,
  protocolReference,
  validateExcelRequest,
} from "./protocol.mjs";

const PACKAGE_ROOT = fileURLToPath(new URL("../..", import.meta.url));
const DEFAULT_STATIC_ROOT = path.join(PACKAGE_ROOT, "apps", "excel-addin", "dist");
const BROWSER_COOKIE = "officekit_excel_browser";
const BROWSER_PANE_HEADER = "x-officekit-pane";
const BROWSER_COOKIE_MAX_AGE_SECONDS = 60 * 60 * 24 * 30;
const POLL_TIMEOUT_MS = 24_000;
const EXECUTION_TIMEOUT_MS = 30_000;
const MAX_EXECUTION_TIMEOUT_MS = 120_000;
const STALE_SESSION_MS = 90_000;
const COMPLETED_REQUEST_RETENTION_MS = 5 * 60_000;
const IDLE_BRIDGE_GRACE_MS = 5 * 60_000;

export async function startExcelBridge({
  paths,
  config,
  secret,
  certificate,
  packageVersion = "0.0.0",
  staticRoot = DEFAULT_STATIC_ROOT,
} = {}) {
  if (!paths || !config || typeof secret !== "string" || !certificate?.leaf?.private || !certificate?.leaf?.cert) {
    throw excelLiveError("invalid-state", "Excel bridge requires initialized state and certificates.");
  }
  const bridge = new ExcelBridge({ paths, config, secret, certificate, packageVersion, staticRoot });
  await bridge.start();
  return bridge;
}

export class ExcelBridge {
  constructor({ paths, config, secret, certificate, packageVersion, staticRoot }) {
    this.paths = paths;
    this.config = config;
    this.secret = secret;
    this.certificate = certificate;
    this.packageVersion = packageVersion;
    this.staticRoot = staticRoot;
    this.origin = `https://localhost:${config.port}`;
    this.sessions = new Map();
    this.sessionByPane = new Map();
    this.servers = [];
    this.lastActivity = Date.now();
  }

  async start() {
    await assertStaticAssets(this.staticRoot);
    const handler = (request, response) => {
      void this.handle(request, response).catch((error) => this.writeError(response, error));
    };
    const credentials = {
      key: this.certificate.leaf.private,
      cert: this.certificate.chain ?? this.certificate.leaf.cert,
    };
    const ipv6 = createServer(credentials, handler);
    try {
      await listen(ipv6, this.config.port, "::1");
      this.servers.push(ipv6);
    } catch (error) {
      ipv6.close();
      if (error?.code !== "EADDRNOTAVAIL" && error?.code !== "EAFNOSUPPORT") throw bridgeListenError(error);
    }
    const ipv4 = createServer(credentials, handler);
    try {
      await listen(ipv4, this.config.port, "127.0.0.1");
      this.servers.push(ipv4);
    } catch (error) {
      ipv4.close();
      await this.close();
      throw bridgeListenError(error);
    }
  }

  async close() {
    for (const session of this.sessions.values()) this.disconnectSession(session, "bridge-stopped");
    const servers = this.servers.splice(0);
    await Promise.all(servers.map((server) => new Promise((resolve) => server.close(resolve))));
  }

  isIdle(now = Date.now()) {
    this.pruneStaleSessions(now);
    // A person needs enough time to upload the generated manifest and click
    // the ribbon after `officekit excel install`. This is still an on-demand
    // process: it owns no login item, service registration, or open socket
    // once the grace window has elapsed without a live workbook.
    return this.sessions.size === 0 && now - this.lastActivity >= IDLE_BRIDGE_GRACE_MS;
  }

  async handle(request, response) {
    this.lastActivity = Date.now();
    this.pruneStaleSessions(this.lastActivity);
    const url = new URL(request.url || "/", this.origin);
    if (request.method === "GET" && STATIC_PATHS.has(url.pathname)) {
      await this.serveStatic(url.pathname, request, response);
      return;
    }
    if (url.pathname === "/v1/browser/bootstrap" && request.method === "POST") {
      this.assertBrowserOrigin(request);
      this.ensureBrowser(request, response);
      this.writeJson(response, 200, {
        protocol: EXCEL_LIVE_PROTOCOL,
        bridge: "ready",
        origin: this.origin,
        limits: protocolReference().limits,
      });
      return;
    }
    if (url.pathname === "/v1/cli/health" && request.method === "GET") {
      this.requireCli(request);
      this.writeJson(response, 200, createExcelSuccess({
        result: {
          bridge: "ready",
          protocol: EXCEL_LIVE_PROTOCOL,
          certificate: this.certificate.certificate,
        },
      }));
      return;
    }
    if (url.pathname === "/v1/cli/sessions" && request.method === "GET") {
      this.requireCli(request);
      this.writeJson(response, 200, createExcelSuccess({ result: { sessions: this.listSessions() } }));
      return;
    }
    if (url.pathname === "/v1/cli/doctor" && request.method === "GET") {
      this.requireCli(request);
      this.writeJson(response, 200, createExcelSuccess({
        result: {
          bridge: "ready",
          origin: this.origin,
          trusted: Boolean(this.config.trusted),
          sessions: this.listSessions(),
          protocol: protocolReference(),
        },
      }));
      return;
    }
    if (url.pathname === "/v1/cli/execute" && request.method === "POST") {
      this.requireCli(request);
      const body = await readJsonBody(request);
      const requestedTimeout = body.timeoutMs ?? EXECUTION_TIMEOUT_MS;
      if (!Number.isSafeInteger(requestedTimeout) || requestedTimeout < 1_000 || requestedTimeout > MAX_EXECUTION_TIMEOUT_MS) {
        throw excelLiveError("invalid-request", `timeoutMs must be from 1000 through ${MAX_EXECUTION_TIMEOUT_MS}.`);
      }
      const result = await this.execute(body.request, requestedTimeout);
      this.writeJson(response, 200, result);
      return;
    }
    if (url.pathname === "/v1/cli/disconnect" && request.method === "POST") {
      this.requireCli(request);
      const body = await readJsonBody(request);
      const session = this.requireSession(body.sessionId);
      this.disconnectSession(session, "cli-disconnect");
      this.writeJson(response, 200, createExcelSuccess({ result: { disconnected: body.sessionId } }));
      return;
    }
    if (url.pathname === "/v1/cli/shutdown" && request.method === "POST") {
      this.requireCli(request);
      this.writeJson(response, 200, createExcelSuccess({ result: { shuttingDown: true } }));
      const timer = setTimeout(() => { void this.close(); }, 10);
      timer.unref();
      return;
    }
    await this.handleBrowserRoute(request, response, url);
  }

  async handleBrowserRoute(request, response, url) {
    const match = /^\/v1\/browser\/sessions\/([^/]+)(?:\/(status|refresh|next|results|disconnect))?$/u.exec(url.pathname);
    if (url.pathname === "/v1/browser/sessions" && request.method === "POST") {
      const browser = this.requireBrowser(request);
      const body = await readJsonBody(request);
      const client = validateBrowserClient(body.client);
      const browserKey = hash(browser);
      const existing = this.sessionByPane.get(browserPaneKey(browserKey, client.paneId));
      if (existing != null && this.sessions.has(existing)) {
        const session = this.sessions.get(existing);
        session.client = client;
        session.lastSeen = Date.now();
        this.writeJson(response, 200, { session: this.publicSession(session) });
        return;
      }
      const session = this.createSession(browserKey, client);
      this.writeJson(response, 201, { session: this.publicSession(session) });
      return;
    }
    if (match == null) throw excelLiveError("not-found", "Excel bridge route was not found.");
    const browser = this.requireBrowser(request);
    const session = this.requireBrowserSession(decodeURIComponent(match[1]), browser, request);
    const action = match[2] ?? "";
    if (action === "status" && request.method === "GET") {
      this.writeJson(response, 200, { connected: true, session: this.publicSession(session) });
      return;
    }
    if (action === "refresh" && request.method === "POST") {
      const body = await readJsonBody(request);
      session.client = validateBrowserClient(body.client);
      session.lastSeen = Date.now();
      this.writeJson(response, 200, { session: this.publicSession(session) });
      return;
    }
    if (action === "next" && request.method === "GET") {
      await this.nextRequest(session, request, response);
      return;
    }
    if (action === "results" && request.method === "POST") {
      const body = await readJsonBody(request, MAX_IMAGE_BYTES + MAX_REQUEST_BYTES);
      const completion = await this.completeRequest(session, body);
      this.writeJson(response, 200, { accepted: true, completion });
      return;
    }
    if (action === "disconnect" && request.method === "POST") {
      this.disconnectSession(session, "browser-disconnect");
      this.writeJson(response, 200, { disconnected: true });
      return;
    }
    throw excelLiveError("not-found", "Excel browser route was not found.");
  }

  async execute(requestValue, timeoutMs) {
    const request = validateExcelRequest(requestValue);
    const session = this.requireSession(request.sessionId);
    const existing = session.byIdempotency.get(request.idempotencyKey);
    const record = existing ?? this.enqueue(session, request);
    try {
      return await waitForResult(record, timeoutMs);
    } catch (error) {
      return createExcelFailure(error, { audit: auditSummary(record) });
    }
  }

  enqueue(session, request) {
    if (session.queue.length + session.inFlight.size >= 32) {
      throw excelLiveError("session-busy", "Excel session queue is full.", { retryable: true });
    }
    const record = {
      requestId: randomUUID(),
      request,
      createdAt: new Date().toISOString(),
      dispatched: false,
      completed: false,
      response: null,
      resolve: null,
      promise: null,
    };
    record.promise = new Promise((resolve) => { record.resolve = resolve; });
    session.queue.push(record);
    session.byIdempotency.set(request.idempotencyKey, record);
    this.wakeSession(session);
    return record;
  }

  async nextRequest(session, request, response) {
    // A shared-runtime task pane keeps its session alive while it is hidden by
    // continuing the long poll. Do not confuse an empty poll with inactivity.
    session.lastSeen = Date.now();
    if (session.inFlight.size > 0) {
      throw excelLiveError(
        "session-busy",
        "Excel session already has an operation in progress.",
        { retryable: true },
      );
    }
    const record = session.queue.shift() ?? await this.waitForNext(session, request, response);
    if (record == null) {
      response.statusCode = 204;
      response.end();
      return;
    }
    record.dispatched = true;
    session.inFlight.set(record.requestId, record);
    session.lastSeen = Date.now();
    this.writeJson(response, 200, {
      requestId: record.requestId,
      request: { operation: record.request.operation, args: record.request.args },
    });
  }

  waitForNext(session, request, response) {
    if (session.waiter != null) {
      throw excelLiveError("session-busy", "Excel session already has an active long-poll request.", { retryable: true });
    }
    return new Promise((resolve) => {
      let timer;
      const cleanup = () => {
        finish(null);
      };
      const finish = (value) => {
        if (session.waiter?.resolve !== finish) return;
        clearTimeout(timer);
        session.waiter = null;
        request.off("aborted", cleanup);
        response.off("close", cleanup);
        resolve(value);
      };
      timer = setTimeout(() => finish(null), POLL_TIMEOUT_MS);
      request.once("aborted", cleanup);
      response.once("close", cleanup);
      session.waiter = { resolve: finish, timer };
    });
  }

  wakeSession(session) {
    if (session.waiter == null || session.queue.length === 0) return;
    const waiter = session.waiter;
    waiter.resolve(session.queue.shift());
  }

  async completeRequest(session, body) {
    if (body == null || typeof body !== "object" || Array.isArray(body) || typeof body.requestId !== "string") {
      throw excelLiveError("invalid-result", "Excel add-in result has an invalid shape.");
    }
    const record = session.inFlight.get(body.requestId);
    if (record == null) throw excelLiveError("unknown-request", "Excel add-in returned an unknown request ID.");
    session.inFlight.delete(record.requestId);
    session.lastSeen = Date.now();
    if (body.ok === true) {
      record.response = createExcelSuccess({ result: body.result ?? {}, audit: auditSummary(record) });
    } else if (body.ok === false && body.error && typeof body.error === "object") {
      record.response = createExcelFailure(excelLiveError(
        typeof body.error.code === "string" ? body.error.code : "office-operation-failed",
        typeof body.error.message === "string" ? body.error.message : "Excel operation failed.",
        { retryable: Boolean(body.error.retryable), maybeApplied: true },
      ), { audit: auditSummary(record) });
    } else {
      throw excelLiveError("invalid-result", "Excel add-in result must contain ok and a result or error.");
    }
    record.completed = true;
    record.resolve(record.response);
    await this.recordAudit(session, record);
    const retentionTimer = setTimeout(() => {
      if (session.byIdempotency.get(record.request.idempotencyKey) === record) {
        session.byIdempotency.delete(record.request.idempotencyKey);
      }
    }, COMPLETED_REQUEST_RETENTION_MS);
    retentionTimer.unref();
    return auditSummary(record);
  }

  async recordAudit(session, record) {
    const audit = auditSummary(record);
    try {
      await appendAuditRecord(this.paths, {
        timestamp: new Date().toISOString(),
        sessionId: session.id,
        workbook: session.client.workbook.name,
        operation: record.request.operation,
        range: rangeSummary(record.request.args),
        requestHash: audit.requestHash,
        status: record.response?.ok ? "ok" : "error",
      });
    } catch {
      // Auditing must not change a completed Excel operation into a failed one.
    }
  }

  createSession(browserKey, client) {
    const session = {
      id: randomUUID(),
      browserKey,
      paneId: client.paneId,
      client,
      createdAt: new Date().toISOString(),
      lastSeen: Date.now(),
      queue: [],
      inFlight: new Map(),
      byIdempotency: new Map(),
      waiter: null,
    };
    this.sessions.set(session.id, session);
    this.sessionByPane.set(browserPaneKey(browserKey, session.paneId), session.id);
    return session;
  }

  disconnectSession(session, reason) {
    if (session.waiter != null) {
      session.waiter.resolve(null);
    }
    const records = [...session.queue, ...session.inFlight.values()];
    session.queue.length = 0;
    session.inFlight.clear();
    for (const record of records) {
      if (record.completed) continue;
      record.completed = true;
      record.response = createExcelFailure(excelLiveError(
        "session-disconnected",
        `Excel session disconnected: ${reason}.`,
        { retryable: true, maybeApplied: record.dispatched },
      ), { audit: auditSummary(record) });
      record.resolve(record.response);
    }
    this.sessions.delete(session.id);
    const paneKey = browserPaneKey(session.browserKey, session.paneId);
    if (this.sessionByPane.get(paneKey) === session.id) {
      this.sessionByPane.delete(paneKey);
    }
  }

  requireSession(id) {
    if (typeof id !== "string") throw excelLiveError("invalid-request", "sessionId is required.");
    const session = this.sessions.get(id);
    if (session == null) throw excelLiveError("session-unavailable", "Excel session is unavailable. Open OfficeKit in the target workbook and connect it.", { retryable: true });
    return session;
  }

  requireBrowserSession(id, browser, request) {
    const session = this.requireSession(id);
    if (session.browserKey !== hash(browser) || session.paneId !== browserPaneId(request)) {
      throw excelLiveError("forbidden", "This browser session cannot access the requested Excel session.");
    }
    return session;
  }

  requireCli(request) {
    const authorization = request.headers.authorization;
    const token = typeof authorization === "string" && authorization.startsWith("Bearer ")
      ? authorization.slice("Bearer ".length)
      : "";
    if (!safeEqual(token, this.secret)) {
      throw excelLiveError("unauthorized", "OfficeKit Excel CLI credentials are invalid.");
    }
  }

  requireBrowser(request) {
    this.assertBrowserOrigin(request);
    const token = parseCookies(request.headers.cookie)[BROWSER_COOKIE];
    if (typeof token !== "string" || !verifyBrowserToken(token, this.secret)) {
      throw excelLiveError("browser-not-paired", "Open the OfficeKit task pane and connect this workbook.");
    }
    return token;
  }

  assertBrowserOrigin(request) {
    const origin = request.headers.origin;
    if (origin != null && origin !== this.origin) {
      throw excelLiveError("forbidden-origin", "Excel bridge rejected a request from another origin.");
    }
  }

  ensureBrowser(request, response) {
    const existing = parseCookies(request.headers.cookie)[BROWSER_COOKIE];
    if (typeof existing === "string" && verifyBrowserToken(existing, this.secret)) return existing;
    const token = createBrowserToken(this.secret);
    response.setHeader("Set-Cookie", `${BROWSER_COOKIE}=${token}; Max-Age=${BROWSER_COOKIE_MAX_AGE_SECONDS}; Path=/v1/browser; Secure; HttpOnly; SameSite=Strict`);
    return token;
  }

  async serveStatic(pathname, request, response) {
    const relative = STATIC_PATHS.get(pathname);
    const content = await readFile(path.join(this.staticRoot, relative));
    if (pathname === "/taskpane.html") this.ensureBrowser(request, response);
    response.statusCode = 200;
    response.setHeader("Content-Type", contentType(relative));
    response.setHeader("Cache-Control", relative.endsWith(".png") ? "public, max-age=86400" : "no-store");
    response.setHeader("X-Content-Type-Options", "nosniff");
    response.setHeader("Referrer-Policy", "no-referrer");
    if (pathname === "/taskpane.html") {
      response.setHeader(
        "Content-Security-Policy",
        "default-src 'self' https://appsforoffice.microsoft.com; connect-src 'self'; img-src 'self' data:; style-src 'self'; script-src 'self' https://appsforoffice.microsoft.com",
      );
    }
    response.end(content);
  }

  writeJson(response, status, value) {
    const body = JSON.stringify(value);
    if (Buffer.byteLength(body) > MAX_IMAGE_BYTES + MAX_REQUEST_BYTES) {
      throw excelLiveError("response-too-large", "Excel bridge response exceeds the safety limit.");
    }
    response.statusCode = status;
    response.setHeader("Content-Type", "application/json; charset=utf-8");
    response.setHeader("Cache-Control", "no-store");
    response.setHeader("X-Content-Type-Options", "nosniff");
    response.end(body);
  }

  writeError(response, error) {
    if (response.writableEnded) return;
    const normalized = toExcelLiveFailure(error);
    const status = statusForError(normalized.code);
    try {
      this.writeJson(response, status, createExcelFailure(error));
    } catch {
      response.statusCode = 500;
      response.end("{\"ok\":false,\"error\":{\"code\":\"internal-error\",\"message\":\"Excel bridge failed.\"}}");
    }
  }

  listSessions() {
    this.pruneStaleSessions(Date.now());
    return [...this.sessions.values()].map((session) => this.publicSession(session));
  }

  publicSession(session) {
    return {
      id: session.id,
      workbook: session.client.workbook,
      capabilities: session.client.capabilities,
      host: session.client.host,
      connectedAt: session.createdAt,
      lastSeenAt: new Date(session.lastSeen).toISOString(),
      queued: session.queue.length + session.inFlight.size,
    };
  }

  pruneStaleSessions(now) {
    for (const session of [...this.sessions.values()]) {
      if (now - session.lastSeen > STALE_SESSION_MS) this.disconnectSession(session, "session-stale");
    }
  }
}

const STATIC_PATHS = new Map([
  ["/", "taskpane.html"],
  ["/taskpane.html", "taskpane.html"],
  ["/taskpane.js", "taskpane.js"],
  ["/taskpane.css", "taskpane.css"],
  ["/support.html", "support.html"],
  ["/assets/officekit-excel-32.png", "assets/officekit-excel-32.png"],
  ["/assets/officekit-excel-80.png", "assets/officekit-excel-80.png"],
]);

async function assertStaticAssets(staticRoot) {
  for (const relative of STATIC_PATHS.values()) {
    try {
      await readFile(path.join(staticRoot, relative));
    } catch {
      throw excelLiveError("addin-assets-missing", "Excel add-in assets are missing. Reinstall OfficeKit or run its package build.");
    }
  }
}

function listen(server, port, host) {
  return new Promise((resolve, reject) => {
    const onError = (error) => {
      server.off("listening", onListening);
      reject(error);
    };
    const onListening = () => {
      server.off("error", onError);
      resolve();
    };
    server.once("error", onError);
    server.once("listening", onListening);
    server.listen({ port, host, exclusive: true });
  });
}

function bridgeListenError(error) {
  if (error?.code === "EADDRINUSE") {
    return excelLiveError("bridge-port-in-use", "OfficeKit Excel bridge port is already in use by another process.");
  }
  return excelLiveError("bridge-start-failed", `OfficeKit Excel bridge could not listen: ${error?.message ?? error}`);
}

async function readJsonBody(request, maximum = MAX_REQUEST_BYTES) {
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    const bytes = Buffer.from(chunk);
    size += bytes.length;
    if (size > maximum) throw excelLiveError("request-too-large", `Excel bridge request exceeds ${maximum} bytes.`);
    chunks.push(bytes);
  }
  try {
    return JSON.parse(Buffer.concat(chunks, size).toString("utf8"));
  } catch (error) {
    throw excelLiveError("invalid-request", `Excel bridge request is not JSON: ${error.message}`);
  }
}

function validateBrowserClient(value) {
  if (value == null || typeof value !== "object" || Array.isArray(value)) {
    throw excelLiveError("invalid-session", "Excel add-in client descriptor must be an object.");
  }
  const paneId = value.paneId;
  if (typeof paneId !== "string" || !/^[A-Za-z0-9_-]{16,128}$/u.test(paneId)) {
    throw excelLiveError("invalid-session", "Excel add-in did not provide a valid task-pane identity.");
  }
  const workbook = value.workbook;
  if (
    workbook == null || typeof workbook !== "object" ||
    typeof workbook.name !== "string" || workbook.name.length === 0 || workbook.name.length > 512 ||
    typeof workbook.activeSheet !== "string" || workbook.activeSheet.length === 0 || workbook.activeSheet.length > 255
  ) {
    throw excelLiveError("invalid-session", "Excel add-in did not report a valid workbook target.");
  }
  const capabilities = value.capabilities != null && typeof value.capabilities === "object" && !Array.isArray(value.capabilities)
    ? Object.fromEntries(Object.entries(value.capabilities).filter(([, capability]) => typeof capability === "boolean"))
    : {};
  const host = value.host != null && typeof value.host === "object" && !Array.isArray(value.host)
    ? {
      platform: typeof value.host.platform === "string" ? value.host.platform.slice(0, 128) : "unknown",
      version: typeof value.host.version === "string" ? value.host.version.slice(0, 128) : "unknown",
      webView: typeof value.host.webView === "string" ? value.host.webView.slice(0, 128) : "unknown",
    }
    : { platform: "unknown", version: "unknown", webView: "unknown" };
  return {
    paneId,
    workbook: { name: workbook.name, activeSheet: workbook.activeSheet },
    capabilities,
    host,
  };
}

function browserPaneId(request) {
  const candidate = request.headers[BROWSER_PANE_HEADER];
  const paneId = Array.isArray(candidate) ? candidate[0] : candidate;
  if (typeof paneId !== "string" || !/^[A-Za-z0-9_-]{16,128}$/u.test(paneId)) {
    throw excelLiveError("forbidden", "Excel bridge rejected a request without the paired task-pane identity.");
  }
  return paneId;
}

function browserPaneKey(browserKey, paneId) {
  return `${browserKey}:${paneId}`;
}

function waitForResult(record, timeoutMs) {
  if (record.completed) return Promise.resolve(record.response);
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(excelLiveError(
        "operation-timeout",
        "Excel did not finish the operation before the timeout. Re-read the target before retrying.",
        { retryable: true, maybeApplied: record.dispatched },
      )), timeoutMs);
    timer.unref();
    record.promise.then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      (error) => {
        clearTimeout(timer);
        reject(error);
      },
    );
  });
}

function auditSummary(record) {
  return {
    requestId: record.requestId,
    sessionId: record.request.sessionId,
    operation: record.request.operation,
    requestHash: hash(JSON.stringify({
      idempotencyKey: record.request.idempotencyKey,
      operation: record.request.operation,
      args: record.request.args,
    })),
  };
}

function rangeSummary(args) {
  const values = [];
  for (const candidate of [args, args?.source, args?.destination]) {
    if (candidate?.sheet && candidate?.range) values.push(`${candidate.sheet}!${candidate.range}`);
  }
  return values.join(",") || undefined;
}

function createBrowserToken(secret) {
  const payload = `${randomUUID()}.${Date.now()}`;
  const signature = createHmac("sha256", secret).update(payload).digest("base64url");
  return `${Buffer.from(payload).toString("base64url")}.${signature}`;
}

function verifyBrowserToken(token, secret) {
  const [encodedPayload, suppliedSignature, extra] = token.split(".");
  if (!encodedPayload || !suppliedSignature || extra) return false;
  let payload;
  try {
    payload = Buffer.from(encodedPayload, "base64url").toString("utf8");
  } catch {
    return false;
  }
  const [identifier, timestamp, tail] = payload.split(".");
  if (!identifier || !timestamp || tail || !/^\d{13}$/u.test(timestamp)) return false;
  if (Date.now() - Number(timestamp) > BROWSER_COOKIE_MAX_AGE_SECONDS * 1000) return false;
  const expected = createHmac("sha256", secret).update(payload).digest("base64url");
  return safeEqual(suppliedSignature, expected);
}

function safeEqual(left, right) {
  const leftBytes = Buffer.from(String(left));
  const rightBytes = Buffer.from(String(right));
  return leftBytes.length === rightBytes.length && timingSafeEqual(leftBytes, rightBytes);
}

function parseCookies(header) {
  if (typeof header !== "string") return {};
  const result = {};
  for (const part of header.split(";")) {
    const separator = part.indexOf("=");
    if (separator < 1) continue;
    result[part.slice(0, separator).trim()] = part.slice(separator + 1).trim();
  }
  return result;
}

function hash(value) {
  return createHash("sha256").update(value).digest("hex");
}

function contentType(relative) {
  if (relative.endsWith(".html")) return "text/html; charset=utf-8";
  if (relative.endsWith(".js")) return "text/javascript; charset=utf-8";
  if (relative.endsWith(".css")) return "text/css; charset=utf-8";
  if (relative.endsWith(".png")) return "image/png";
  return "application/octet-stream";
}

function statusForError(code) {
  if (["unauthorized", "browser-not-paired"].includes(code)) return 401;
  if (["forbidden", "forbidden-origin"].includes(code)) return 403;
  if (["not-found", "session-unavailable"].includes(code)) return 404;
  if (["session-busy", "bridge-port-in-use"].includes(code)) return 409;
  if (["request-too-large", "response-too-large"].includes(code)) return 413;
  if (["operation-timeout"].includes(code)) return 504;
  return 400;
}
