import assert from "node:assert/strict";
import https from "node:https";
import net from "node:net";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { lstat, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";

import { startPowerPointBridge } from "../src/live/bridge.mjs";
import { createPowerPointLiveAdapter, POWERPOINT_LIVE_OPERATIONS, validatePowerPointRequest } from "../src/live/adapters/powerpoint.mjs";
import { OfficeLiveError } from "../src/live/errors.mjs";
import { renderPowerPointManifest } from "../src/powerpoint-live/manifest.mjs";
import {
  ensureExcelCertificates,
  generateCertificateBundle,
  persistCertificateMetadata,
} from "../src/excel-live/certificates.mjs";
import {
  initializePowerPointConfiguration,
  readPowerPointConfiguration,
  resolvePowerPointStatePaths,
} from "../src/powerpoint-live/state.mjs";
import { bridgeRequest } from "../src/powerpoint-live/client.mjs";

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
const primaryPaneId = "ppt-pane-primary-0001";
const temporary = await mkdtemp(path.join(os.tmpdir(), "officekit-powerpoint-live-"));
let bridge;

try {
  const manifest = renderPowerPointManifest({
    addinId: "8f3e6f45-8f0a-4f25-93a2-55a6b71f7f23",
    port: 47214,
    packageVersion: "0.7.0",
  });
  assert.match(manifest, /<Host Name="Presentation"\/>/);
  assert.match(manifest, /PowerPointApi" MinVersion="1\.8"/);
  assert.match(manifest, /SharedRuntime" MinVersion="1\.1"/);
  assert.match(manifest, /ReadWriteDocument/);
  assert.match(manifest, /https:\/\/localhost:47214\/powerpoint\/taskpane\.html/);
  assert.match(manifest, /lifetime="long"/);

  const baseRequest = {
    protocol: 1,
    sessionId: "ppt-session-1234",
    idempotencyKey: "ppt-idempotency-1234",
  };
  for (const [operation, args] of Object.entries({
    read_presentation: {},
    read_slides: {},
    read_slide: { slideId: "slide-1" },
    read_selection: {},
    write_text: { slideId: "slide-1", shapeId: "shape-1", text: "Updated", expectedSnapshot: { id: "shape-1", text: "Original", left: 10 } },
    add_textbox: { slideId: "slide-1", text: "New", left: 10, top: 10, width: 100, height: 50 },
    add_shape: { slideId: "slide-1", type: "rect", left: 10, top: 10, width: 100, height: 50 },
    add_image: { slideId: "slide-1", imageData: "data:image/png;base64,AA==", left: 10, top: 10, width: 100, height: 50 },
    update_shape: { slideId: "slide-1", shapeId: "shape-1", left: 20 },
    delete_shape: { slideId: "slide-1", shapeId: "shape-1" },
    add_slide: { layoutId: "layout-1" },
    read_slide_image: { slideId: "slide-1", width: 640, height: 360 },
    save: {},
  })) {
    assert.equal(validatePowerPointRequest({ ...baseRequest, operation, args }).operation, operation);
  }
  assert.deepEqual(POWERPOINT_LIVE_OPERATIONS, [
    "read_presentation", "read_slides", "read_slide", "read_selection", "write_text",
    "add_textbox", "add_shape", "add_image", "update_shape", "delete_shape", "add_slide",
    "read_slide_image", "save",
  ]);
  assert.throws(
    () => validatePowerPointRequest({ ...baseRequest, operation: "run_officejs", args: {} }),
    (error) => error instanceof OfficeLiveError && error.code === "unsupported-operation",
  );
  assert.throws(
    () => validatePowerPointRequest({ ...baseRequest, operation: "update_shape", args: { slideId: "slide-1", shapeId: "shape-1" } }),
    /requires a geometry or text change/,
  );
  assert.throws(
    () => validatePowerPointRequest({ ...baseRequest, operation: "add_image", args: { slideId: "slide-1", imageData: "data:image/png;base64,AA==", left: 0, top: 0, width: 100, height: 100, sourcePath: "/tmp/a.png" } }),
    /not supported/,
  );
  assert.throws(
    () => validatePowerPointRequest({ ...baseRequest, operation: "add_image", args: { slideId: "slide-1", imageData: `data:image/svg+xml;base64,${Buffer.from("<svg><script>alert(1)</script></svg>").toString("base64")}`, left: 0, top: 0, width: 100, height: 100 } }),
    /unsafe SVG/,
  );
  assert.throws(
    () => validatePowerPointRequest({ ...baseRequest, host: "excel", operation: "save", args: {} }),
    (error) => error instanceof OfficeLiveError && error.code === "forbidden",
  );

  const certificateFixture = await generateCertificateBundle(new Date("2026-01-01T00:00:00.000Z"));
  assert.match(certificateFixture.root.cert, /BEGIN CERTIFICATE/);
  const paths = resolvePowerPointStatePaths({ env: { OFFICEKIT_POWERPOINT_HOME: path.join(temporary, "state") } });
  let state = await initializePowerPointConfiguration(paths, { port: await unusedPort() });
  assert.equal(state.config.application, "powerpoint");
  assert.equal(state.config.addinId, "8f3e6f45-8f0a-4f25-93a2-55a6b71f7f23");
  const generated = await ensureExcelCertificates(paths, state.config);
  await persistCertificateMetadata(paths, generated.certificate);
  state = await readPowerPointConfiguration(paths);
  const certificates = await ensureExcelCertificates(paths, state.config);
  const adapter = createPowerPointLiveAdapter({ staticRoot: path.join(repoRoot, "apps", "powerpoint-addin", "dist") });
  bridge = await startPowerPointBridge({
    paths,
    ...state,
    certificate: certificates,
    packageVersion: "0.7.0",
    adapter,
  });

  const page = await browserRequest(state, "GET", "/powerpoint/taskpane.html");
  assert.equal(page.status, 200);
  assert.match(page.text, /PowerPoint/);
  assert.match(page.cookie, /officekit_powerpoint_browser=/);
  const bootstrap = await browserRequest(state, "POST", "/v1/browser/bootstrap", {}, page.cookie);
  assert.equal(bootstrap.status, 200);
  assert.equal(bootstrap.json.protocol, 1);
  const connected = await browserRequest(state, "POST", "/v1/browser/sessions", {
    client: {
      paneId: primaryPaneId,
      presentation: { name: "Unsaved deck", activeSlideId: "slide-1", slideCount: 2 },
      capabilities: { powerpointApi18: true, sharedRuntime: true, slideImage: true },
      host: { platform: "PC", version: "16.0", webView: "WebView2" },
    },
  }, page.cookie);
  assert.equal(connected.status, 201);
  const session = connected.json.session;
  assert.match(session.id, /^powerpoint-/);
  assert.equal(session.application, "powerpoint");
  assert.equal(session.presentation.name, "Unsaved deck");
  assert.equal(session.host.version, "16.0");
  const switchedTarget = await browserRequest(state, "POST", "/v1/browser/sessions", {
    client: {
      paneId: primaryPaneId,
      presentation: { name: "Another deck" },
      capabilities: {},
      host: { platform: "PC", version: "16.0" },
    },
  }, page.cookie);
  assert.equal(switchedTarget.status, 409, JSON.stringify(switchedTarget));
  assert.equal(switchedTarget.json.error.code, "target-changed");

  const listed = await bridgeRequest(state, "GET", "/v1/cli/sessions");
  assert.equal(listed.ok, true);
  assert.equal(listed.result.sessions[0].id, session.id);

  const request = { ...baseRequest, sessionId: session.id, idempotencyKey: "ppt-save-12345678", operation: "save", args: {} };
  const parkedPoll = browserRequest(state, "GET", `/v1/browser/sessions/${session.id}/next`, undefined, page.cookie);
  const completionPromise = bridgeRequest(state, "POST", "/v1/cli/execute", { request });
  const next = await parkedPoll;
  assert.equal(next.status, 200);
  assert.equal(next.json.request.operation, "save");
  const resultResponse = await browserRequest(state, "POST", `/v1/browser/sessions/${session.id}/results`, {
    requestId: next.json.requestId,
    ok: true,
    result: { saveRequested: true },
  }, page.cookie);
  assert.equal(resultResponse.status, 200);
  const completed = await completionPromise;
  assert.equal(completed.ok, true);
  assert.equal(completed.result.saveRequested, true);
  assert.equal((await bridgeRequest(state, "POST", "/v1/cli/execute", { request })).audit.operation, "save");

  const secondary = await browserRequest(state, "POST", "/v1/browser/sessions", {
    client: {
      paneId: "ppt-pane-secondary-0002",
      presentation: { name: "Second deck" },
      capabilities: {},
      host: { platform: "PC", version: "16.0" },
    },
  }, page.cookie, undefined, "ppt-pane-secondary-0002");
  assert.notEqual(secondary.json.session.id, session.id);
  const crossPane = await browserRequest(state, "GET", `/v1/browser/sessions/${session.id}/status`, undefined, page.cookie, undefined, "ppt-pane-secondary-0002");
  assert.equal(crossPane.status, 403, JSON.stringify(crossPane));

  const addinJs = await readFile(path.join(repoRoot, "apps", "powerpoint-addin", "dist", "taskpane.js"), "utf8");
  const addinHtml = await readFile(path.join(repoRoot, "apps", "powerpoint-addin", "dist", "taskpane.html"), "utf8");
  assert.match(addinHtml, /appsforoffice\.microsoft\.com\/lib\/1\/hosted\/office\.js/);
  for (const operation of POWERPOINT_LIVE_OPERATIONS) assert.match(addinJs, new RegExp(operation));
  assert.doesNotMatch(addinJs, /run_officejs/);

  const mocked = await runMockedTaskpane({ repoRoot, state, cookie: page.cookie });
  const mockSession = (await bridgeRequest(state, "GET", "/v1/cli/sessions")).result.sessions.find((candidate) => candidate.presentation?.name === "Mock live deck");
  assert.ok(mockSession);
  const writeRequest = {
    protocol: 1,
    sessionId: mockSession.id,
    idempotencyKey: "mock-write-text-1234",
    operation: "write_text",
    args: { slideId: "slide-1", shapeId: "shape-1", text: "Changed", expectedSnapshot: { id: "shape-1", text: "Original", left: 10, width: 300 } },
  };
  const writeResult = bridgeRequest(state, "POST", "/v1/cli/execute", { request: writeRequest, timeoutMs: 1_500 });
  await waitFor(() => mocked.deck.shapes[0].text === "Changed", `PowerPoint typed text write (${mocked.fetchPaths.join(" | ")})`);
  const writeOutcome = await writeResult;
  assert.equal(writeOutcome.ok, true, `${JSON.stringify(writeOutcome)} paths=${mocked.fetchPaths.join(" | ")}`);
  const imageRequest = {
    protocol: 1,
    sessionId: mockSession.id,
    idempotencyKey: "mock-slide-image-1234",
    operation: "read_slide_image",
    args: { slideId: "slide-1", width: 640, height: 360 },
  };
  const imageResult = bridgeRequest(state, "POST", "/v1/cli/execute", { request: imageRequest, timeoutMs: 1_500 });
  await waitFor(() => mocked.imageCalls === 1, "PowerPoint slide image");
  assert.equal((await imageResult).result.mimeType, "image/png");
  assert.match((await imageResult).result.data, /^data:image\/png;base64,/);
  const saveRequest = { ...writeRequest, idempotencyKey: "mock-save-1234", operation: "save", args: {} };
  const saveResult = bridgeRequest(state, "POST", "/v1/cli/execute", { request: saveRequest, timeoutMs: 1_500 });
  await waitFor(() => mocked.saveCalls === 1, "PowerPoint explicit save");
  assert.equal((await saveResult).result.saveRequested, true);
  await mocked.disconnect();

  console.log("powerpoint live smoke ok");
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

function browserRequest(state, method, requestPath, body, cookie, origin = `https://localhost:${state.config.port}`, paneId = primaryPaneId) {
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
        ...(serialized == null ? {} : { "content-type": "application/json", "content-length": Buffer.byteLength(serialized) }),
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
  const original = Object.fromEntries(["Office", "PowerPoint", "document", "fetch"].map((key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)]));
  const elements = new Map();
  const deck = {
    title: "Mock live deck",
    shapes: [{ id: "shape-1", name: "Title", type: "textBox", text: "Original", left: 10, top: 10, width: 300, height: 80 }],
    slides: [{ id: "slide-1" }],
    selectedSlideId: "slide-1",
  };
  let currentCookie = cookie;
  let saveCalls = 0;
  let imageCalls = 0;
  const fetchPaths = [];
  const element = (id) => {
    if (!elements.has(id)) elements.set(id, new FakeElement());
    return elements.get(id);
  };
  setGlobal("document", { getElementById: element });
  setGlobal("Office", {
    HostType: { PowerPoint: "PowerPoint" },
    context: {
      platform: "PC",
      diagnostics: { version: "16.0" },
      requirements: { isSetSupported: () => true },
    },
    onReady(callback) { void callback({ host: "PowerPoint" }); },
  });
  setGlobal("PowerPoint", {
    GeometricShapeType: { rectangle: "rect", ellipse: "ellipse", roundRectangle: "roundRect", triangle: "triangle", hexagon: "hexagon", diamond: "diamond" },
    ConnectorType: { straight: "straight" },
    run(callback) {
      const context = createMockContext(deck, {
        onSave: () => { saveCalls += 1; },
        onImage: () => { imageCalls += 1; },
      });
      return callback(context);
    },
  });
  setGlobal("fetch", async (url, options = {}) => {
    fetchPaths.push(`${options.method ?? "GET"} ${String(url)} pane=${options.headers?.["x-officekit-pane"] ?? "none"}`);
    const response = await browserRequest(state, options.method ?? "GET", String(url), options.body === undefined ? undefined : JSON.parse(options.body), currentCookie, undefined, options.headers?.["x-officekit-pane"] ?? primaryPaneId);
    if (response.cookie !== undefined) currentCookie = response.cookie;
    if (response.status >= 300) fetchPaths.push(`response ${response.status} ${response.text}`);
    return { status: response.status, ok: response.status >= 200 && response.status < 300, json: async () => response.json ?? {} };
  });
  const taskpaneUrl = `${pathToFileURL(path.join(repoRoot, "apps", "powerpoint-addin", "dist", "taskpane.js")).href}?mock=${Date.now()}-${Math.random()}`;
  try {
    await import(taskpaneUrl);
    await waitFor(() => element("connect").disabled === false, "PowerPoint task pane readiness");
    element("connect").trigger("click");
    await waitFor(async () => (await bridgeRequest(state, "GET", "/v1/cli/sessions")).result.sessions.some((session) => session.presentation?.name === "Mock live deck"), `PowerPoint task pane connection (${element("setup-copy").textContent})`);
    return {
      deck,
      get saveCalls() { return saveCalls; },
      get imageCalls() { return imageCalls; },
      fetchPaths,
      async disconnect() {
        element("disconnect").trigger("click");
        await waitFor(async () => !(await bridgeRequest(state, "GET", "/v1/cli/sessions")).result.sessions.some((session) => session.presentation?.name === "Mock live deck"), "PowerPoint task pane disconnect");
        restoreGlobals(original);
      },
    };
  } catch (error) {
    restoreGlobals(original);
    throw error;
  }
}

function createMockContext(deck, { onSave, onImage }) {
  const makeShape = (raw) => {
    const shape = {
      id: raw.id,
      name: raw.name,
      type: raw.type,
      left: raw.left,
      top: raw.top,
      width: raw.width,
      height: raw.height,
      textFrame: {
        textRange: {
          get text() { return raw.text ?? ""; },
          set text(value) { raw.text = value; },
        },
      },
      load() {},
      delete() { deck.shapes = deck.shapes.filter((candidate) => candidate.id !== raw.id); },
    };
    return shape;
  };
  const slide = {
    id: deck.slides[0].id,
    load() {},
    shapes: {
      get items() { return deck.shapes.map(makeShape); },
      load() {},
      getItem(id) {
        const raw = deck.shapes.find((candidate) => candidate.id === id);
        if (!raw) throw new Error(`shape ${id} not found`);
        return makeShape(raw);
      },
      addTextBox(text, geometry) {
        const raw = { id: `shape-${deck.shapes.length + 1}`, type: "textBox", text, ...geometry };
        deck.shapes.push(raw);
        return makeShape(raw);
      },
      addGeometricShape(type, geometry) {
        const raw = { id: `shape-${deck.shapes.length + 1}`, type, text: "", ...geometry };
        deck.shapes.push(raw);
        return makeShape(raw);
      },
      addLine(type, geometry) {
        const raw = { id: `shape-${deck.shapes.length + 1}`, type, text: "", ...geometry };
        deck.shapes.push(raw);
        return makeShape(raw);
      },
      addPicture(data, geometry) {
        const raw = { id: `shape-${deck.shapes.length + 1}`, type: "image", text: "", imageData: data };
        Object.assign(raw, geometry);
        deck.shapes.push(raw);
        return makeShape(raw);
      },
    },
    getImageAsBase64() {
      onImage();
      return { value: "iVBORw0KGgo=" };
    },
  };
  const slides = {
    get items() { return deck.slides.map(() => slide); },
    load() {},
    getItem(id) { if (id !== slide.id) throw new Error(`slide ${id} not found`); return slide; },
    add() {
      const raw = { id: `slide-${deck.slides.length + 1}` };
      deck.slides.push(raw);
      return slide;
    },
  };
  const selectedSlides = { items: [slide], load() {} };
  return {
    presentation: {
      title: deck.title,
      load() {},
      slides,
      getSelectedSlides: () => selectedSlides,
      getSelectedShapes: () => ({ items: deck.shapes.map(makeShape).slice(0, 1), load() {} }),
      save: onSave,
    },
    sync: async () => {},
  };
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
  Object.defineProperty(globalThis, key, { value, configurable: true, enumerable: true, writable: true });
}
