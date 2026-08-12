import assert from "node:assert/strict";
import https from "node:https";
import net from "node:net";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { lstat, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";

import { startExcelBridge } from "../src/excel-live/bridge.mjs";
import { bridgeRequest } from "../src/excel-live/client.mjs";
import {
  ensureExcelCertificates,
  generateCertificateBundle,
  persistCertificateMetadata,
} from "../src/excel-live/certificates.mjs";
import { ExcelLiveError } from "../src/excel-live/errors.mjs";
import { renderExcelManifest } from "../src/excel-live/manifest.mjs";
import {
  createExcelFailure,
  EXCEL_LIVE_OPERATIONS,
  validateExcelRequest,
} from "../src/excel-live/protocol.mjs";
import {
  initializeExcelConfiguration,
  readExcelConfiguration,
  resolveExcelStatePaths,
  updateExcelConfiguration,
} from "../src/excel-live/state.mjs";
import { doctorExcel, installExcel, runExcelCommand } from "../src/excel-live/cli.mjs";

class FakeElement {
  constructor() {
    this.hidden = false;
    this.disabled = false;
    this.textContent = "";
    this.listeners = new Map();
  }

  addEventListener(event, callback) {
    this.listeners.set(event, callback);
  }

  trigger(event) {
    this.listeners.get(event)?.();
  }
}

const repoRoot = path.resolve(import.meta.dirname, "..");
const primaryPaneId = "test-pane-primary-0001";
const temporary = await mkdtemp(path.join(os.tmpdir(), "officekit-excel-live-"));
let bridge;

try {
  const certificateFixture = await generateCertificateBundle(new Date("2026-01-01T00:00:00.000Z"));
  assert.match(certificateFixture.root.cert, /BEGIN CERTIFICATE/);
  assert.match(certificateFixture.leaf.cert, /BEGIN CERTIFICATE/);

  const manifest = renderExcelManifest({
    addinId: "d209533c-4ca9-4aa1-b64b-467bbdd23fc0",
    port: 47213,
    packageVersion: "0.6.0",
  });
  assert.match(manifest, /https:\/\/localhost:47213\/taskpane\.html/);
  assert.match(manifest, /ReadWriteDocument/);
  assert.match(manifest, /SharedRuntime/);
  assert.match(manifest, /lifetime="long"/);

  const request = validateExcelRequest({
    protocol: 1,
    sessionId: "session-1234",
    idempotencyKey: "idempotency-1234",
    operation: "write_range",
    args: {
      sheet: "Sheet1",
      range: "A1:B2",
      values: [[1, 2], [3, 4]],
      numberFormat: [["0", "0"], ["0", "0"]],
    },
  });
  assert.equal(request.operation, "write_range");
  assert.equal(EXCEL_LIVE_OPERATIONS.includes("save"), true);
  assert.throws(
    () => validateExcelRequest({ ...request, operation: "run_officejs" }),
    (error) => error instanceof ExcelLiveError && error.code === "unsupported-operation",
  );
  assert.throws(
    () => validateExcelRequest({
      ...request,
      args: { sheet: "Sheet1", range: "A1:B2", values: [[1, 2]], formulas: [["=1"], ["=2"]] },
    }),
    /matching rectangular dimensions/,
  );
  assert.throws(
    () => validateExcelRequest({
      ...request,
      operation: "search_workbook",
      args: { query: "revenue", options: { useRegex: true } },
    }),
    /not supported/,
  );
  assert.throws(
    () => validateExcelRequest({
      ...request,
      operation: "format_range",
      args: { sheet: "Sheet1", range: "A1", format: { formula: "=NOW()" } },
    }),
    /not supported/,
  );
  assert.equal(createExcelFailure(new ExcelLiveError("test", "expected")).error.code, "test");
  for (const [operation, args] of Object.entries({
    read_ranges: { sheet: "Sheet1", ranges: ["A1"] },
    search_workbook: { query: "revenue" },
    list_items: { kind: "tables" },
    write_range: { sheet: "Sheet1", range: "A1", values: [[1]] },
    clear_range: { sheet: "Sheet1", range: "A1", applyTo: "contents" },
    update_sheet: { action: "add", name: "Plan" },
    update_workbook: { calculationMode: "Automatic" },
    copy_range_to: { source: { sheet: "Sheet1", range: "A1" }, destination: { sheet: "Sheet1", range: "B1" } },
    read_range_image: { sheet: "Sheet1", range: "A1:B2" },
    read_sheets_metadata: {},
    resize_range: { sheet: "Sheet1", range: "A:A", autofitColumns: true },
    update_sheet_view: { sheet: "Sheet1", freezeRows: 1 },
    format_range: { sheet: "Sheet1", range: "A1", format: { fill: { color: "#ffffff" } } },
    chart: { action: "create", sheet: "Sheet1", name: "Revenue", type: "ColumnClustered", sourceRange: "A1:B5" },
    table: { action: "create", sheet: "Sheet1", name: "RevenueTable", range: "A1:B5" },
    pivot_table: { action: "create", sheet: "Sheet1", name: "RevenuePivot", source: "RevenueTable", destination: "D1" },
    save: {},
  })) {
    assert.equal(validateExcelRequest({
      protocol: 1,
      sessionId: "session-1234",
      idempotencyKey: `operation-${operation}`,
      operation,
      args,
    }).operation, operation);
  }

  const port = await unusedPort();
  const paths = resolveExcelStatePaths({ env: { OFFICEKIT_EXCEL_HOME: path.join(temporary, "state") } });
  let state = await initializeExcelConfiguration(paths, { port });
  const generated = await ensureExcelCertificates(paths, state.config);
  await assert.rejects(lstat(paths.rootKey), { code: "ENOENT" }, "the trusted root private key must not remain on disk");
  await persistCertificateMetadata(paths, generated.certificate);
  state = await readExcelConfiguration(paths);
  await writeFile(paths.rootKey, "legacy development root key\n", { mode: 0o600 });
  const certificates = await ensureExcelCertificates(paths, state.config);
  await assert.rejects(lstat(paths.rootKey), { code: "ENOENT" }, "a legacy root private key must be removed during normal startup");
  bridge = await startExcelBridge({
    paths,
    ...state,
    certificate: certificates,
    packageVersion: "0.6.0",
  });

  const page = await browserRequest(state, "GET", "/taskpane.html");
  assert.equal(page.status, 200);
  assert.match(page.text, /OfficeKit/);
  assert.match(page.cookie, /officekit_excel_browser=/);
  const bootstrap = await browserRequest(state, "POST", "/v1/browser/bootstrap", {}, page.cookie);
  assert.equal(bootstrap.status, 200);
  assert.equal(bootstrap.json.protocol, 1);
  assert.equal(bootstrap.cookie, undefined, "a valid browser cookie must be retained, not rotated");

  const connected = await browserRequest(
    state,
    "POST",
    "/v1/browser/sessions",
    {
      client: {
        paneId: primaryPaneId,
        workbook: { name: "Unsaved workbook", activeSheet: "Sheet1" },
        capabilities: { excelApi18: true, sharedRuntime: true },
        host: { platform: "OfficeOnline", version: "16.0" },
      },
    },
    page.cookie,
  );
  const session = connected.json.session;
  assert.equal(session.workbook.name, "Unsaved workbook");

  const listed = await bridgeRequest(state, "GET", "/v1/cli/sessions");
  assert.equal(listed.ok, true);
  assert.equal(listed.result.sessions.length, 1);
  assert.equal(listed.result.sessions[0].id, session.id);

  const diagnosed = await doctorExcel({
    statePaths: paths,
    platform: "darwin",
    probeTrust: async () => ({ trusted: true }),
    ensureBridge: async () => {},
  });
  assert.equal(diagnosed.result.host.status, "ready");
  assert.equal(diagnosed.result.host.sessions[0].host.version, "16.0");
  assert.match(diagnosed.result.host.required.desktopWindowForZoom, /ExcelApiDesktop/);

  const excelRequest = {
    protocol: 1,
    sessionId: session.id,
    idempotencyKey: "save-12345678",
    operation: "save",
    args: {},
  };
  let parkedPollReturned = false;
  const parkedPoll = browserRequest(state, "GET", `/v1/browser/sessions/${session.id}/next`, undefined, page.cookie)
    .then((value) => {
      parkedPollReturned = true;
      return value;
    });
  await new Promise((resolve) => setTimeout(resolve, 80));
  assert.equal(parkedPollReturned, false, "the Add-in long poll must wait instead of spinning empty requests");
  const completionPromise = bridgeRequest(state, "POST", "/v1/cli/execute", { request: excelRequest });
  const next = await parkedPoll;
  assert.equal(next.json.request.operation, "save");
  const concurrentPoll = await browserRequest(state, "GET", `/v1/browser/sessions/${session.id}/next`, undefined, page.cookie);
  assert.equal(concurrentPoll.status, 409, "a session may only dispatch one Excel operation at a time");
  assert.equal(concurrentPoll.json.error.code, "session-busy");
  const completion = await browserRequest(
    state,
    "POST",
    `/v1/browser/sessions/${session.id}/results`,
    { requestId: next.json.requestId, ok: true, result: { saved: true } },
    page.cookie,
  );
  assert.equal(completion.json.accepted, true);
  const executed = await completionPromise;
  assert.equal(executed.ok, true);
  assert.equal(executed.result.saved, true);
  assert.equal(executed.audit.operation, "save");
  assert.doesNotMatch(JSON.stringify(executed.audit), /Unsave|Sheet1/);

  const idempotent = await bridgeRequest(state, "POST", "/v1/cli/execute", { request: excelRequest });
  assert.deepEqual(idempotent, executed, "same idempotency key must return the original completion");

  const forbidden = await browserRequest(
    state,
    "POST",
    "/v1/browser/sessions",
    { client: { workbook: { name: "Other", activeSheet: "Sheet1" } } },
    page.cookie,
    "https://example.invalid",
  );
  assert.equal(forbidden.status, 403);
  assert.equal(forbidden.json.error.code, "forbidden-origin");

  const audit = await readFile(paths.audit, "utf8");
  assert.match(audit, /"operation":"save"/);
  assert.doesNotMatch(audit, /"saved":true/);

  const secondaryPaneId = "test-pane-secondary-0002";
  const secondWorkbook = await browserRequest(
    state,
    "POST",
    "/v1/browser/sessions",
    {
      client: {
        paneId: secondaryPaneId,
        workbook: { name: "Second workbook", activeSheet: "Plan" },
        capabilities: { excelApi18: true, sharedRuntime: true },
        host: { platform: "PC", version: "16.0" },
      },
    },
    page.cookie,
    undefined,
    secondaryPaneId,
  );
  assert.notEqual(secondWorkbook.json.session.id, session.id, "two workbook task panes must receive distinct sessions");
  const crossPane = await browserRequest(
    state,
    "GET",
    `/v1/browser/sessions/${session.id}/status`,
    undefined,
    page.cookie,
    undefined,
    secondaryPaneId,
  );
  assert.equal(crossPane.status, 403, "one task pane must not operate another workbook session");

  const mockAddin = await runMockedTaskpane({
    repoRoot,
    state,
    cookie: page.cookie,
    bridge,
  });
  const taskpaneSession = (await bridgeRequest(state, "GET", "/v1/cli/sessions"))
    .result.sessions.find((candidate) => candidate.workbook.name === "Mock live workbook");
  assert.ok(taskpaneSession, "compiled task pane must register its open workbook through the bridge");
  const mockExecution = bridgeRequest(state, "POST", "/v1/cli/execute", {
    request: {
      protocol: 1,
      sessionId: taskpaneSession.id,
      idempotencyKey: "mock-taskpane-save-1234",
      operation: "save",
      args: {},
    },
  });
  await waitFor(() => mockAddin.saveCalls === 1, "compiled task pane save");
  const mockResult = await mockExecution;
  assert.equal(mockResult.ok, true);
  assert.deepEqual(mockAddin.saveBehaviors, ["Prompt"], "an unsaved workbook must leave Save As path choice to Excel");
  assert.equal(mockAddin.requirementChecks.some(([set, version]) => set === "ExcelApi" && version === "1.11"), true);
  await waitFor(() => /"operation": "save"/.test(mockAddin.element("audit-output").textContent), "compiled task pane audit rendering");
  assert.match(mockAddin.element("audit-output").textContent, /"operation": "save"/);
  await mockAddin.disconnect();

  const installPaths = resolveExcelStatePaths({ env: { OFFICEKIT_EXCEL_HOME: path.join(temporary, "install") } });
  const installation = await installExcel({
    statePaths: installPaths,
    packageVersion: "0.6.0",
    confirmed: true,
    trust: async (candidatePaths, config) => (await updateExcelConfiguration(candidatePaths, (current) => ({ ...current, trusted: true }))).config,
    ensureBridge: async () => {},
  });
  assert.equal(installation.ok, true);
  assert.match(await readFile(installation.result.manifest, "utf8"), /OfficeKitTaskpane/);

  const recoveryPaths = resolveExcelStatePaths({ env: { OFFICEKIT_EXCEL_HOME: path.join(temporary, "recovery") } });
  let recoveryState = await initializeExcelConfiguration(recoveryPaths, { port: await unusedPort() });
  const initialRecoveryCertificate = await ensureExcelCertificates(recoveryPaths, recoveryState.config);
  await persistCertificateMetadata(recoveryPaths, initialRecoveryCertificate.certificate);
  await rm(recoveryPaths.leafCertificate, { force: true });
  await installExcel({
    statePaths: recoveryPaths,
    packageVersion: "0.6.0",
    confirmed: true,
    trust: async (candidatePaths) => (await updateExcelConfiguration(candidatePaths, (current) => ({ ...current, trusted: true }))).config,
    ensureBridge: async () => {},
  });
  recoveryState = await readExcelConfiguration(recoveryPaths);
  assert.notEqual(
    recoveryState.config.certificate.leafFingerprintSha256,
    initialRecoveryCertificate.certificate.leafFingerprintSha256,
    "a regenerated certificate set must replace stale metadata before the bridge is started",
  );

  await assert.rejects(
    runExcelCommand(["sessions"], { platform: "linux" }),
    (error) => error instanceof ExcelLiveError && error.code === "unsupported-platform",
  );
  await assert.rejects(
    runExcelCommand(["execute", "-"], { platform: "darwin" }),
    (error) => error instanceof ExcelLiveError && error.code === "invalid-request-file",
  );

  const addinJs = await readFile(path.join(repoRoot, "apps", "excel-addin", "dist", "taskpane.js"), "utf8");
  const addinHtml = await readFile(path.join(repoRoot, "apps", "excel-addin", "dist", "taskpane.html"), "utf8");
  assert.match(addinHtml, /appsforoffice\.microsoft\.com\/lib\/1\/hosted\/office\.js/);
  assert.match(addinJs, /read_ranges/);
  assert.match(addinJs, /pivot_table/);
  assert.doesNotMatch(addinJs, /run_officejs/);

  console.log("excel live smoke ok");
} finally {
  await bridge?.close();
  await rm(temporary, { recursive: true, force: true });
}

function unusedPort() {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.once("error", reject);
    server.listen({ host: "127.0.0.1", port: 0 }, () => {
      const address = server.address();
      server.close((error) => error ? reject(error) : resolve(address.port));
    });
  });
}

function browserRequest(
  state,
  method,
  requestPath,
  body,
  cookie,
  origin = `https://localhost:${state.config.port}`,
  paneId = primaryPaneId,
) {
  return new Promise((resolve, reject) => {
    const serialized = body === undefined ? null : JSON.stringify(body);
    const request = https.request({
      host: "localhost",
      port: state.config.port,
      method,
      path: requestPath,
      rejectUnauthorized: false,
      headers: {
        ...(cookie === undefined ? {} : { cookie }),
        ...(requestPath.startsWith("/v1/browser") ? { origin, "x-officekit-pane": paneId } : {}),
        ...(serialized == null ? {} : {
          "content-type": "application/json",
          "content-length": Buffer.byteLength(serialized),
        }),
      },
    }, (response) => {
      const chunks = [];
      response.on("data", (chunk) => chunks.push(Buffer.from(chunk)));
      response.once("error", reject);
      response.once("end", () => {
        const text = Buffer.concat(chunks).toString("utf8");
        const contentType = String(response.headers["content-type"] ?? "");
        resolve({
          status: response.statusCode,
          text,
          json: contentType.includes("application/json") && text ? JSON.parse(text) : null,
          cookie: response.headers["set-cookie"]?.[0]?.split(";", 1)[0],
        });
      });
    });
    request.once("error", reject);
    request.end(serialized ?? undefined);
  });
}

async function runMockedTaskpane({ repoRoot, state, cookie }) {
  const original = {
    Office: Object.getOwnPropertyDescriptor(globalThis, "Office"),
    Excel: Object.getOwnPropertyDescriptor(globalThis, "Excel"),
    document: Object.getOwnPropertyDescriptor(globalThis, "document"),
    fetch: Object.getOwnPropertyDescriptor(globalThis, "fetch"),
  };
  const elements = new Map();
  const requirementChecks = [];
  let saveCalls = 0;
  const saveBehaviors = [];
  const element = (id) => {
    if (!elements.has(id)) elements.set(id, new FakeElement());
    return elements.get(id);
  };
  let currentCookie = cookie;
  setGlobal("document", { getElementById: element });
  setGlobal("Office", {
    HostType: { Excel: "Excel" },
    context: {
      platform: "PC",
      diagnostics: { version: "16.0" },
      requirements: {
        isSetSupported(set, version) {
          requirementChecks.push([set, version]);
          return set === "ExcelApi" || set === "SharedRuntime";
        },
      },
    },
    onReady(callback) {
      void callback({ host: "Excel" });
    },
  });
  setGlobal("Excel", {
    SaveBehavior: { prompt: "Prompt" },
    run(callback) {
      const activeSheet = { name: "Sheet1", load() {} };
      const workbook = {
        name: "Mock live workbook",
        load() {},
        save(behavior) { saveCalls += 1; saveBehaviors.push(behavior); },
        worksheets: { getActiveWorksheet: () => activeSheet },
      };
      return callback({ workbook, sync: async () => {} });
    },
  });
  setGlobal("fetch", async (url, options = {}) => {
    const body = options.body === undefined ? undefined : JSON.parse(options.body);
    const paneId = options.headers?.["x-officekit-pane"];
    const response = await browserRequest(
      state,
      options.method ?? "GET",
      String(url),
      body,
      currentCookie,
      undefined,
      typeof paneId === "string" ? paneId : primaryPaneId,
    );
    if (response.cookie !== undefined) currentCookie = response.cookie;
    return {
      status: response.status,
      ok: response.status >= 200 && response.status < 300,
      json: async () => response.json ?? {},
    };
  });
  const taskpaneUrl = `${pathToFileURL(path.join(repoRoot, "apps", "excel-addin", "dist", "taskpane.js")).href}?mock=${Date.now()}-${Math.random()}`;
  try {
    await import(taskpaneUrl);
    await waitFor(() => element("connect").disabled === false, "compiled task pane readiness");
    element("connect").trigger("click");
    await waitFor(async () => {
      const sessions = await bridgeRequest(state, "GET", "/v1/cli/sessions");
      return sessions.result.sessions.some((session) => session.workbook.name === "Mock live workbook");
    }, "compiled task pane connection");
    return {
      get saveCalls() { return saveCalls; },
      saveBehaviors,
      requirementChecks,
      element,
      async disconnect() {
        element("disconnect").trigger("click");
        await waitFor(async () => {
          const sessions = await bridgeRequest(state, "GET", "/v1/cli/sessions");
          return !sessions.result.sessions.some((session) => session.workbook.name === "Mock live workbook");
        }, "compiled task pane disconnect");
        restoreGlobals(original);
      },
    };
  } catch (error) {
    restoreGlobals(original);
    throw error;
  }
}

async function waitFor(predicate, label, timeoutMs = 5_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (await predicate()) return;
    await new Promise((resolve) => setTimeout(resolve, 25));
  }
  throw new Error(`Timed out waiting for ${label}.`);
}

function restoreGlobals(original) {
  for (const [key, descriptor] of Object.entries(original)) {
    if (descriptor === undefined) delete globalThis[key];
    else Object.defineProperty(globalThis, key, descriptor);
  }
}

function setGlobal(key, value) {
  Object.defineProperty(globalThis, key, {
    value,
    configurable: true,
    enumerable: true,
    writable: true,
  });
}
