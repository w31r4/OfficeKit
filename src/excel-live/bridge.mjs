import { createHash, createHmac, randomUUID, timingSafeEqual } from "node:crypto";
import { createServer } from "node:https";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { appendAuditRecord } from "./state.mjs";
import { excelLiveError, toExcelLiveFailure } from "./errors.mjs";
import { createExcelLiveAdapter } from "../live/adapters/excel.mjs";
import {
  EXCEL_LIVE_PROTOCOL,
  MAX_IMAGE_BYTES,
  MAX_REQUEST_BYTES,
} from "./protocol.mjs";

const PACKAGE_ROOT = fileURLToPath(new URL("../..", import.meta.url));
const DEFAULT_STATIC_ROOT = path.join(PACKAGE_ROOT, "apps", "excel-addin", "dist");
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
  adapter = createExcelLiveAdapter({ staticRoot }),
} = {}) {
  if (!paths || !config || typeof secret !== "string" || !certificate?.leaf?.private || !certificate?.leaf?.cert) {
    throw adapter.error("invalid-state", `${adapter.targetLabel ?? adapter.host} bridge requires initialized state and certificates.`);
  }
  const bridge = new ExcelBridge({ paths, config, secret, certificate, packageVersion, staticRoot, adapter });
  await bridge.start();
  return bridge;
}

export class ExcelBridge {
  constructor({ paths, config, secret, certificate, packageVersion, staticRoot, adapter = createExcelLiveAdapter({ staticRoot }) }) {
    this.paths = paths;
    this.config = config;
    this.secret = secret;
    this.certificate = certificate;
    this.packageVersion = packageVersion;
    this.adapter = adapter;
    this.staticRoot = adapter.staticRoot ?? staticRoot;
    this.origin = `https://localhost:${config.port}`;
    this.sessions = new Map();
    this.sessionByPane = new Map();
    this.servers = [];
    this.lastActivity = Date.now();
  }

  async start() {
    await assertStaticAssets(this.staticRoot, this.adapter);
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
      if (error?.code !== "EADDRNOTAVAIL" && error?.code !== "EAFNOSUPPORT") throw bridgeListenError(error, this.adapter);
    }
    const ipv4 = createServer(credentials, handler);
    try {
      await listen(ipv4, this.config.port, "127.0.0.1");
      this.servers.push(ipv4);
    } catch (error) {
      ipv4.close();
      await this.close();
      throw bridgeListenError(error, this.adapter);
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
    if (request.method === "GET" && this.adapter.staticPaths.has(url.pathname)) {
      await this.serveStatic(url.pathname, request, response);
      return;
    }
    if (url.pathname === "/v1/browser/bootstrap" && request.method === "POST") {
      this.assertBrowserOrigin(request);
      this.ensureBrowser(request, response);
      this.writeJson(response, 200, {
        protocol: this.adapter.protocol().protocol ?? EXCEL_LIVE_PROTOCOL,
        bridge: "ready",
        origin: this.origin,
        limits: this.adapter.protocol().limits,
      });
      return;
    }
    if (url.pathname === "/v1/cli/health" && request.method === "GET") {
      this.requireCli(request);
      this.writeJson(response, 200, this.adapter.success({
        result: {
          bridge: "ready",
          protocol: this.adapter.protocol().protocol ?? EXCEL_LIVE_PROTOCOL,
          certificate: this.certificate.certificate,
        },
      }));
      return;
    }
    if (url.pathname === "/v1/cli/sessions" && request.method === "GET") {
      this.requireCli(request);
      this.writeJson(response, 200, this.adapter.success({ result: { sessions: this.listSessions() } }));
      return;
    }
    if (url.pathname === "/v1/cli/doctor" && request.method === "GET") {
      this.requireCli(request);
      this.writeJson(response, 200, this.adapter.success({
        result: {
          bridge: "ready",
          origin: this.origin,
          trusted: Boolean(this.config.trusted),
          sessions: this.listSessions(),
          protocol: this.adapter.protocol(),
        },
      }));
      return;
    }
    if (url.pathname === "/v1/cli/execute" && request.method === "POST") {
      this.requireCli(request);
      const body = await readJsonBody(request, this.requestLimit(), this.adapter);
      const requestedTimeout = body.timeoutMs ?? EXECUTION_TIMEOUT_MS;
      if (!Number.isSafeInteger(requestedTimeout) || requestedTimeout < 1_000 || requestedTimeout > MAX_EXECUTION_TIMEOUT_MS) {
        throw this.adapter.error("invalid-request", `timeoutMs must be from 1000 through ${MAX_EXECUTION_TIMEOUT_MS}.`);
      }
      const result = await this.execute(body.request, requestedTimeout);
      this.writeJson(response, 200, result);
      return;
    }
    if (url.pathname === "/v1/cli/disconnect" && request.method === "POST") {
      this.requireCli(request);
      const body = await readJsonBody(request, this.requestLimit(), this.adapter);
      const session = this.requireSession(body.sessionId);
      this.disconnectSession(session, "cli-disconnect");
      this.writeJson(response, 200, this.adapter.success({ result: { disconnected: body.sessionId } }));
      return;
    }
    if (url.pathname === "/v1/cli/shutdown" && request.method === "POST") {
      this.requireCli(request);
      this.writeJson(response, 200, this.adapter.success({ result: { shuttingDown: true } }));
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
      const body = await readJsonBody(request, this.requestLimit(), this.adapter);
      const client = this.adapter.validateClient(body.client);
      const browserKey = hash(browser);
      const existing = this.sessionByPane.get(browserPaneKey(browserKey, client.paneId));
      if (existing != null && this.sessions.has(existing)) {
        const session = this.sessions.get(existing);
        if (typeof this.adapter.sameTarget === "function" && !this.adapter.sameTarget(session.client, client)) {
          throw this.adapter.error("target-changed", `This task pane is already paired with another ${this.adapter.targetLabel}. Disconnect it before connecting a different document.`);
        }
        session.client = client;
        session.lastSeen = Date.now();
        this.writeJson(response, 200, { session: this.publicSession(session) });
        return;
      }
      const session = this.createSession(browserKey, client);
      this.writeJson(response, 201, { session: this.publicSession(session) });
      return;
    }
    if (match == null) throw this.adapter.error("not-found", this.adapter.routeNotFoundMessage);
    const browser = this.requireBrowser(request);
    const session = this.requireBrowserSession(decodeURIComponent(match[1]), browser, request);
    const action = match[2] ?? "";
    if (action === "status" && request.method === "GET") {
      this.writeJson(response, 200, { connected: true, session: this.publicSession(session) });
      return;
    }
    if (action === "refresh" && request.method === "POST") {
      const body = await readJsonBody(request, this.requestLimit(), this.adapter);
      session.client = this.adapter.validateClient(body.client);
      session.lastSeen = Date.now();
      this.writeJson(response, 200, { session: this.publicSession(session) });
      return;
    }
    if (action === "next" && request.method === "GET") {
      await this.nextRequest(session, request, response);
      return;
    }
    if (action === "results" && request.method === "POST") {
      const body = await readJsonBody(request, this.responseLimit(), this.adapter);
      const completion = await this.completeRequest(session, body);
      this.writeJson(response, 200, { accepted: true, completion });
      return;
    }
    if (action === "disconnect" && request.method === "POST") {
      this.disconnectSession(session, "browser-disconnect");
      this.writeJson(response, 200, { disconnected: true });
      return;
    }
    throw this.adapter.error("not-found", this.adapter.routeNotFoundMessage);
  }

  async execute(requestValue, timeoutMs) {
    const request = this.adapter.validateRequest(requestValue);
    const session = this.requireSession(request.sessionId);
    const existing = session.byIdempotency.get(request.idempotencyKey);
    const record = existing ?? this.enqueue(session, request);
    try {
      return await waitForResult(record, timeoutMs, this.adapter);
    } catch (error) {
      return this.adapter.failure(error, { audit: auditSummary(record) });
    }
  }

  enqueue(session, request) {
    if (session.queue.length + session.inFlight.size >= 32) {
      throw this.adapter.error("session-busy", `${this.adapter.targetLabel} session queue is full.`, { retryable: true });
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
      throw this.adapter.error(
        "session-busy",
        `${this.adapter.targetLabel} session already has an operation in progress.`,
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
      throw this.adapter.error("session-busy", `${this.adapter.targetLabel} session already has an active long-poll request.`, { retryable: true });
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
      throw this.adapter.error("invalid-result", `${this.adapter.targetLabel} add-in result has an invalid shape.`);
    }
    const record = session.inFlight.get(body.requestId);
    if (record == null) throw this.adapter.error("unknown-request", `${this.adapter.targetLabel} add-in returned an unknown request ID.`);
    session.inFlight.delete(record.requestId);
    session.lastSeen = Date.now();
    if (body.ok === true) {
      record.response = this.adapter.success({ result: body.result ?? {}, audit: auditSummary(record) });
    } else if (body.ok === false && body.error && typeof body.error === "object") {
      record.response = this.adapter.failure(this.adapter.error(
        typeof body.error.code === "string" ? body.error.code : "office-operation-failed",
        typeof body.error.message === "string" ? body.error.message : this.adapter.operationFailureMessage,
        { retryable: Boolean(body.error.retryable), maybeApplied: true },
      ), { audit: auditSummary(record) });
    } else {
      throw this.adapter.error("invalid-result", `${this.adapter.targetLabel} add-in result must contain ok and a result or error.`);
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
        ...this.adapter.audit({ session, record, summary: audit }),
      });
    } catch {
      // Auditing must not change a completed Excel operation into a failed one.
    }
  }

  createSession(browserKey, client) {
    const session = {
      id: this.adapter.sessionIdPrefix ? `${this.adapter.sessionIdPrefix}-${randomUUID()}` : randomUUID(),
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
      record.response = this.adapter.failure(this.adapter.error(
        "session-disconnected",
        this.adapter.disconnectedMessage(reason),
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
    if (typeof id !== "string") throw this.adapter.error("invalid-request", "sessionId is required.");
    const session = this.sessions.get(id);
    if (session == null) throw this.adapter.error("session-unavailable", this.adapter.unavailableMessage, { retryable: true });
    return session;
  }

  requireBrowserSession(id, browser, request) {
    const session = this.requireSession(id);
    if (session.browserKey !== hash(browser) || session.paneId !== browserPaneId(request, this.adapter)) {
      throw this.adapter.error("forbidden", `This browser session cannot access the requested ${this.adapter.targetLabel}.`);
    }
    return session;
  }

  requireCli(request) {
    const authorization = request.headers.authorization;
    const token = typeof authorization === "string" && authorization.startsWith("Bearer ")
      ? authorization.slice("Bearer ".length)
      : "";
    if (!safeEqual(token, this.secret)) {
      throw this.adapter.error("unauthorized", `OfficeKit ${this.adapter.targetLabel} CLI credentials are invalid.`);
    }
  }

  requireBrowser(request) {
    this.assertBrowserOrigin(request);
    const token = parseCookies(request.headers.cookie)[this.adapter.browserCookie];
    if (typeof token !== "string" || !verifyBrowserToken(token, this.secret)) {
      throw this.adapter.error("browser-not-paired", `Open the OfficeKit task pane and connect this ${this.adapter.targetLabel}.`);
    }
    return token;
  }

  assertBrowserOrigin(request) {
    const origin = request.headers.origin;
    if (origin != null && origin !== this.origin) {
      throw this.adapter.error("forbidden-origin", `${this.adapter.targetLabel} bridge rejected a request from another origin.`);
    }
  }

  ensureBrowser(request, response) {
    const existing = parseCookies(request.headers.cookie)[this.adapter.browserCookie];
    if (typeof existing === "string" && verifyBrowserToken(existing, this.secret)) return existing;
    const token = createBrowserToken(this.secret);
    response.setHeader("Set-Cookie", `${this.adapter.browserCookie}=${token}; Max-Age=${BROWSER_COOKIE_MAX_AGE_SECONDS}; Path=/v1/browser; Secure; HttpOnly; SameSite=Strict`);
    return token;
  }

  async serveStatic(pathname, request, response) {
    const relative = this.adapter.staticPaths.get(pathname);
    const content = await readFile(path.join(this.staticRoot, relative));
    if (relative === "taskpane.html") this.ensureBrowser(request, response);
    response.statusCode = 200;
    response.setHeader("Content-Type", contentType(relative));
    response.setHeader("Cache-Control", relative.endsWith(".png") ? "public, max-age=86400" : "no-store");
    response.setHeader("X-Content-Type-Options", "nosniff");
    response.setHeader("Referrer-Policy", "no-referrer");
    if (relative === "taskpane.html") {
      response.setHeader(
        "Content-Security-Policy",
        "default-src 'self' https://appsforoffice.microsoft.com; connect-src 'self'; img-src 'self' data:; style-src 'self'; script-src 'self' https://appsforoffice.microsoft.com",
      );
    }
    response.end(content);
  }

  writeJson(response, status, value) {
    const body = JSON.stringify(value);
    if (Buffer.byteLength(body) > this.responseLimit()) {
      throw this.adapter.error("response-too-large", `${this.adapter.targetLabel} bridge response exceeds the safety limit.`);
    }
    response.statusCode = status;
    response.setHeader("Content-Type", "application/json; charset=utf-8");
    response.setHeader("Cache-Control", "no-store");
    response.setHeader("X-Content-Type-Options", "nosniff");
    response.end(body);
  }

  writeError(response, error) {
    if (response.writableEnded) return;
    const normalized = error?.code ? error : toExcelLiveFailure(error);
    const status = statusForError(normalized.code);
    try {
      this.writeJson(response, status, this.adapter.failure(error));
    } catch {
      response.statusCode = 500;
      response.end(`{"ok":false,"error":{"code":"internal-error","message":"${this.adapter.targetLabel} bridge failed."}}`);
    }
  }

  listSessions() {
    this.pruneStaleSessions(Date.now());
    return [...this.sessions.values()].map((session) => this.publicSession(session));
  }

  publicSession(session) {
    return {
      id: session.id,
      application: this.adapter.host,
      ...this.adapter.describeClient(session.client),
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

  requestLimit() {
    return this.adapter.protocol().limits?.maxRequestBytes ?? MAX_REQUEST_BYTES;
  }

  responseLimit() {
    return this.adapter.protocol().limits?.maxResponseBytes ?? (MAX_IMAGE_BYTES + MAX_REQUEST_BYTES);
  }
}

async function assertStaticAssets(staticRoot, adapter = createExcelLiveAdapter({ staticRoot })) {
  for (const relative of adapter.staticPaths.values()) {
    try {
      await readFile(path.join(staticRoot, relative));
    } catch {
      throw adapter.error("addin-assets-missing", adapter.assetErrorMessage);
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

function bridgeListenError(error, adapter = createExcelLiveAdapter()) {
  if (error?.code === "EADDRINUSE") {
    return adapter.error("bridge-port-in-use", `OfficeKit ${adapter.targetLabel} bridge port is already in use by another process.`);
  }
  return adapter.error("bridge-start-failed", `OfficeKit ${adapter.targetLabel} bridge could not listen: ${error?.message ?? error}`);
}

async function readJsonBody(request, maximum = MAX_REQUEST_BYTES, adapter = createExcelLiveAdapter()) {
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    const bytes = Buffer.from(chunk);
    size += bytes.length;
    if (size > maximum) throw adapter.error("request-too-large", `${adapter.targetLabel} bridge request exceeds ${maximum} bytes.`);
    chunks.push(bytes);
  }
  try {
    return JSON.parse(Buffer.concat(chunks, size).toString("utf8"));
  } catch (error) {
    throw adapter.error("invalid-request", `${adapter.targetLabel} bridge request is not JSON: ${error.message}`);
  }
}

function browserPaneId(request, adapter = createExcelLiveAdapter()) {
  const candidate = request.headers[adapter.browserPaneHeader ?? "x-officekit-pane"];
  const paneId = Array.isArray(candidate) ? candidate[0] : candidate;
  if (typeof paneId !== "string" || !/^[A-Za-z0-9_-]{16,128}$/u.test(paneId)) {
    throw adapter.error("forbidden", `${adapter.targetLabel} bridge rejected a request without the paired task-pane identity.`);
  }
  return paneId;
}

function browserPaneKey(browserKey, paneId) {
  return `${browserKey}:${paneId}`;
}

function waitForResult(record, timeoutMs, adapter = createExcelLiveAdapter()) {
  if (record.completed) return Promise.resolve(record.response);
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(excelLiveError(
        "operation-timeout",
        `${adapter.targetLabel} did not finish the operation before the timeout. Re-read the target before retrying.`,
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
  if (["session-busy", "bridge-port-in-use", "target-changed"].includes(code)) return 409;
  if (["request-too-large", "response-too-large"].includes(code)) return 413;
  if (["operation-timeout"].includes(code)) return 504;
  return 400;
}
