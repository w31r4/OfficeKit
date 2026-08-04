/* global Office, PowerPoint */

const paneId = taskPaneIdentity();
let activeSession = null;
let polling = false;
const maxResultBytes = 9_000_000;

const connectButton = byId("connect");
const disconnectButton = byId("disconnect");
const refreshButton = byId("refresh");
const setup = byId("setup");
const setupCopy = byId("setup-copy");
const sessionCard = byId("session");
const presentationName = byId("presentation-name");
const sessionId = byId("session-id");
const connectionStatus = byId("connection-status");
const diagnosticsOutput = byId("diagnostics-output");
const auditOutput = byId("audit-output");

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    setSetup("Open OfficeKit from Microsoft PowerPoint.", false);
    return;
  }
  try {
    const bootstrap = await requestJson("/v1/browser/bootstrap", undefined, "POST");
    showDiagnostics(bootstrap);
    connectButton.disabled = false;
    setSetup("Click Connect OfficeKit to share this open presentation with your Agent.", true);
  } catch (error) {
    setSetup(`OfficeKit bridge is unavailable: ${messageOf(error)}`, false);
    showDiagnostics({ bridge: "unavailable", error: messageOf(error) });
  }
});

connectButton.addEventListener("click", () => void connect());
disconnectButton.addEventListener("click", () => void disconnect());
refreshButton.addEventListener("click", () => void refresh());

async function connect() {
  connectButton.disabled = true;
  try {
    const client = await presentationDescriptor();
    const response = await requestJson("/v1/browser/sessions", { client }, "POST");
    activeSession = response.session;
    showSession(activeSession);
    void poll(activeSession.id);
  } catch (error) {
    setSetup(`Could not connect: ${messageOf(error)}`, true);
  } finally {
    connectButton.disabled = false;
  }
}

async function refresh() {
  if (activeSession == null) return;
  try {
    const response = await requestJson(`/v1/browser/sessions/${encodeURIComponent(activeSession.id)}/refresh`, { client: await presentationDescriptor() }, "POST");
    activeSession = response.session;
    showSession(activeSession);
  } catch (error) {
    connectionStatus.textContent = `Refresh failed: ${messageOf(error)}`;
  }
}

async function disconnect() {
  if (activeSession == null) return;
  try { await requestJson(`/v1/browser/sessions/${encodeURIComponent(activeSession.id)}/disconnect`, {}, "POST"); } catch { /* bridge may already be gone */ }
  activeSession = null;
  polling = false;
  sessionCard.hidden = true;
  setup.hidden = false;
  setSetup("Disconnected. Connect again to share this presentation.", true);
}

async function poll(id) {
  if (polling) return;
  polling = true;
  while (activeSession?.id === id) {
    try {
      const response = await fetch(`/v1/browser/sessions/${encodeURIComponent(id)}/next`, { credentials: "same-origin", cache: "no-store", headers: { "x-officekit-pane": paneId } });
      if (response.status === 204) continue;
      if (response.status === 404 || response.status === 410) { await disconnect(); break; }
      const envelope = await parseResponse(response);
      connectionStatus.textContent = `Running ${envelope.request.operation}…`;
      const result = await executeOperation(envelope.request);
      const accepted = await requestJson(`/v1/browser/sessions/${encodeURIComponent(id)}/results`, { requestId: envelope.requestId, ...result }, "POST");
      showAudit(accepted.completion);
      connectionStatus.textContent = "Waiting for OfficeKit commands";
    } catch (error) {
      connectionStatus.textContent = `Reconnecting: ${messageOf(error)}`;
      await delay(900);
    }
  }
  polling = false;
}

async function presentationDescriptor() {
  return PowerPoint.run(async (context) => {
    const presentation = context.presentation;
    presentation.load("title");
    const selectedSlides = presentation.getSelectedSlides();
    selectedSlides.load("items/id");
    await context.sync();
    return {
      host: "powerpoint",
      paneId,
      presentation: {
        name: typeof presentation.title === "string" && presentation.title ? presentation.title : "Unsaved presentation",
        ...(typeof presentation.id === "string" && presentation.id ? { id: presentation.id } : {}),
        ...(selectedSlides.items[0]?.id ? { activeSlideId: selectedSlides.items[0].id } : {}),
      },
      capabilities: {
        powerpointApi15: Office.context.requirements.isSetSupported("PowerPointApi", "1.5"),
        powerpointApi18: Office.context.requirements.isSetSupported("PowerPointApi", "1.8"),
        sharedRuntime: Office.context.requirements.isSetSupported("SharedRuntime", "1.1"),
        slideImage: Office.context.requirements.isSetSupported("PowerPointApi", "1.8"),
        save: Office.context.requirements.isSetSupported("PowerPointApi", "1.3"),
      },
      host: { platform: Office.context.platform, version: Office.context.diagnostics?.version ?? "unknown", webView: detectedWebView() },
    };
  });
}

async function executeOperation(request) {
  try {
    requireOperationCapability(request.operation);
    const result = await PowerPoint.run(async (context) => executeWithContext(context, request.operation, request.args));
    if (jsonByteLength(result) > maxResultBytes) throw capabilityError("response-too-large", "PowerPoint operation result exceeds the bridge limit.");
    return { ok: true, result: result ?? {} };
  } catch (error) {
    const code = classifyOfficeError(error);
    return { ok: false, error: { code, message: messageOf(error), retryable: !["unsupported-capability", "response-too-large", "stale-target", "object-not-found"].includes(code) } };
  }
}

function requireOperationCapability(operation) {
  const version = ["read_slide_image", "add_image"].includes(operation) ? "1.8" : "1.5";
  if (!Office.context.requirements.isSetSupported("PowerPointApi", version)) throw capabilityError("unsupported-capability", `${operation} requires PowerPointApi ${version}.`);
}

async function executeWithContext(context, operation, args) {
  switch (operation) {
    case "read_presentation": return readPresentation(context);
    case "read_slides": return readSlides(context, args);
    case "read_slide": return readSlide(context, args.slideId);
    case "read_selection": return readSelection(context);
    case "write_text": return writeText(context, args);
    case "add_textbox": return addTextBox(context, args);
    case "add_shape": return addShape(context, args);
    case "add_image": return addImage(context, args);
    case "update_shape": return updateShape(context, args);
    case "delete_shape": return deleteShape(context, args);
    case "add_slide": return addSlide(context, args);
    case "read_slide_image": return readSlideImage(context, args);
    case "save": return savePresentation(context);
    default: throw capabilityError("unsupported-capability", `Unsupported PowerPoint operation: ${operation}`);
  }
}

async function readPresentation(context) {
  const presentation = context.presentation;
  presentation.load("title");
  const slides = presentation.slides;
  slides.load("items/id");
  await context.sync();
  return { name: presentation.title || "Unsaved presentation", slideCount: slides.items.length, slideIds: slides.items.map((slide) => slide.id) };
}

async function readSlides(context, args) {
  const presentation = context.presentation;
  const slides = args.slideIds?.length
    ? args.slideIds.map((id) => presentation.slides.getItem(id))
    : null;
  if (slides) for (const slide of slides) loadSlide(slide);
  else presentation.slides.load("items/id");
  await context.sync();
  return { slides: (slides ?? presentation.slides.items).map(slideSummary) };
}

async function readSlide(context, slideId) {
  const slide = context.presentation.slides.getItem(slideId);
  loadSlide(slide);
  await context.sync();
  return { slide: slideSummary(slide) };
}

async function readSelection(context) {
  const slides = context.presentation.getSelectedSlides();
  const shapes = context.presentation.getSelectedShapes();
  slides.load("items/id");
  shapes.load("items/id,name,type,left,top,width,height");
  await context.sync();
  return { slideIds: slides.items.map((slide) => slide.id), shapes: shapes.items.map(shapeSummary) };
}

async function writeText(context, args) {
  const shape = context.presentation.slides.getItem(args.slideId).shapes.getItem(args.shapeId);
  shape.load(args.expectedSnapshot === undefined ? "id,textFrame/textRange/text" : "id,name,type,left,top,width,height,textFrame/textRange/text");
  await context.sync();
  const currentText = shape.textFrame?.textRange?.text ?? "";
  if (args.expectedText !== undefined && currentText !== args.expectedText) throw capabilityError("stale-target", "The target shape text changed before this request.");
  assertExpectedSnapshot(shape, args.expectedSnapshot);
  shape.textFrame.textRange.text = args.text;
  await context.sync();
  return { slideId: args.slideId, shapeId: args.shapeId, text: args.text };
}

async function addTextBox(context, args) {
  const slide = context.presentation.slides.getItem(args.slideId);
  const shape = slide.shapes.addTextBox(args.text, { left: args.left, top: args.top, width: args.width, height: args.height });
  shape.load("id");
  await context.sync();
  return { slideId: args.slideId, shapeId: shape.id };
}

async function addShape(context, args) {
  const slide = context.presentation.slides.getItem(args.slideId);
  const options = { left: args.left, top: args.top, width: args.width, height: args.height };
  const shape = args.type === "line"
    ? slide.shapes.addLine(PowerPoint.ConnectorType.straight, options)
    : slide.shapes.addGeometricShape(shapeType(args.type), options);
  if (args.text !== undefined) shape.textFrame.textRange.text = args.text;
  shape.load("id");
  await context.sync();
  return { slideId: args.slideId, shapeId: shape.id };
}

async function addImage(context, args) {
  const slide = context.presentation.slides.getItem(args.slideId);
  if (typeof slide.shapes.addPicture !== "function") throw capabilityError("unsupported-capability", "This PowerPoint build does not expose typed picture insertion.");
  const shape = slide.shapes.addPicture(args.imageData.replace(/^data:image\/(?:png|jpeg|gif|svg\+xml);base64,/u, ""), {
    left: args.left, top: args.top, width: args.width, height: args.height,
  });
  if (args.altText !== undefined) shape.altTextDescription = args.altText;
  shape.load("id");
  await context.sync();
  return { slideId: args.slideId, shapeId: shape.id };
}

async function updateShape(context, args) {
  const shape = context.presentation.slides.getItem(args.slideId).shapes.getItem(args.shapeId);
  shape.load(args.expectedSnapshot === undefined ? "id,textFrame/textRange/text" : "id,name,type,left,top,width,height,textFrame/textRange/text");
  await context.sync();
  const currentText = shape.textFrame?.textRange?.text ?? "";
  if (args.expectedText !== undefined && currentText !== args.expectedText) throw capabilityError("stale-target", "The target shape text changed before this request.");
  assertExpectedSnapshot(shape, args.expectedSnapshot);
  for (const key of ["left", "top", "width", "height"]) if (args[key] !== undefined) shape[key] = args[key];
  if (args.text !== undefined) shape.textFrame.textRange.text = args.text;
  await context.sync();
  return { slideId: args.slideId, shapeId: args.shapeId, updated: true };
}

async function deleteShape(context, args) {
  const shape = context.presentation.slides.getItem(args.slideId).shapes.getItem(args.shapeId);
  if (args.expectedSnapshot !== undefined) {
    shape.load("id,name,type,left,top,width,height,textFrame/textRange/text");
    await context.sync();
    assertExpectedSnapshot(shape, args.expectedSnapshot);
  }
  shape.delete();
  await context.sync();
  return { slideId: args.slideId, shapeId: args.shapeId, deleted: true };
}

async function addSlide(context, args) {
  const options = {
    ...(args.slideMasterId === undefined ? {} : { slideMasterId: args.slideMasterId }),
    ...(args.layoutId === undefined && args.layout === undefined ? {} : { layoutId: args.layoutId ?? args.layout }),
  };
  const slide = context.presentation.slides.add(Object.keys(options).length === 0 ? undefined : options);
  slide.load("id");
  await context.sync();
  return { slideId: slide.id };
}

async function readSlideImage(context, args) {
  const slide = context.presentation.slides.getItem(args.slideId);
  const result = slide.getImageAsBase64({ width: args.width, height: args.height });
  await context.sync();
  return { slideId: args.slideId, mimeType: "image/png", data: `data:image/png;base64,${result.value}` };
}

async function savePresentation(context) {
  context.presentation.save();
  await context.sync();
  return { saveRequested: true };
}

function loadSlide(slide) {
  slide.load("id");
  slide.shapes.load("items/id,name,type,left,top,width,height,textFrame/textRange/text");
}

function slideSummary(slide) {
  return { id: slide.id, shapes: slide.shapes?.items?.map(shapeSummary) ?? [] };
}

function shapeSummary(shape) {
  const text = shape.textFrame?.textRange?.text;
  return { id: shape.id, ...(shape.name ? { name: shape.name } : {}), ...(shape.type ? { type: shape.type } : {}), ...(typeof text === "string" ? { text } : {}), left: shape.left, top: shape.top, width: shape.width, height: shape.height };
}

function assertExpectedSnapshot(shape, expected) {
  if (expected === undefined) return;
  const actual = shapeSummary(shape);
  for (const key of ["id", "name", "type", "text"]) {
    if (expected[key] !== undefined && actual[key] !== expected[key]) {
      throw capabilityError("stale-target", `The target shape ${key} changed before this request.`);
    }
  }
  for (const key of ["left", "top", "width", "height"]) {
    if (expected[key] !== undefined && (typeof actual[key] !== "number" || Math.abs(actual[key] - expected[key]) > 0.01)) {
      throw capabilityError("stale-target", `The target shape ${key} changed before this request.`);
    }
  }
}

function shapeType(type) {
  const names = { rect: "rectangle", ellipse: "ellipse", roundRect: "roundRectangle", triangle: "triangle", hexagon: "hexagon", diamond: "diamond" };
  return PowerPoint.GeometricShapeType?.[names[type]] ?? names[type];
}

async function requestJson(pathname, body, method = "GET") {
  const response = await fetch(pathname, {
    method,
    credentials: "same-origin",
    cache: "no-store",
    headers: {
      "x-officekit-pane": paneId,
      ...(body === undefined ? {} : { "content-type": "application/json" }),
    },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  return parseResponse(response);
}

async function parseResponse(response) {
  const body = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(body?.error?.message ?? `OfficeKit bridge returned HTTP ${response.status}.`);
  return body;
}

function showSession(session) { setup.hidden = true; sessionCard.hidden = false; presentationName.textContent = session.presentation?.name ?? "—"; sessionId.textContent = session.id; }
function showAudit(value) { auditOutput.textContent = JSON.stringify(value ?? {}, null, 2); }
function showDiagnostics(value) { diagnosticsOutput.textContent = JSON.stringify(value ?? {}, null, 2); }
function setSetup(value, showButton) { setup.hidden = false; setupCopy.textContent = value; connectButton.hidden = !showButton; }
function byId(id) { return document.getElementById(id); }
function taskPaneIdentity() { const bytes = new Uint8Array(16); crypto.getRandomValues(bytes); return `ppt-${Array.from(bytes, (value) => value.toString(16).padStart(2, "0")).join("")}`; }
function detectedWebView() { return typeof navigator === "object" ? navigator.userAgent.slice(0, 128) : "unknown"; }
function jsonByteLength(value) { return new TextEncoder().encode(JSON.stringify(value ?? {})).length; }
function capabilityError(code, message) { const error = new Error(message); error.code = code; return error; }
function classifyOfficeError(error) {
  if (["unsupported-capability", "response-too-large", "stale-target", "object-not-found"].includes(error?.code)) return error.code;
  const message = messageOf(error);
  if (error?.code === "ItemNotFound" || /(?:item|object|shape|slide)[^\n]*(?:not found|does not exist)/iu.test(message)) return "object-not-found";
  return "office-operation-failed";
}
function messageOf(error) { return error instanceof Error ? error.message : String(error); }
function delay(milliseconds) { return new Promise((resolve) => setTimeout(resolve, milliseconds)); }
