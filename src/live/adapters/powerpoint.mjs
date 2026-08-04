import path from "node:path";
import { fileURLToPath } from "node:url";

import { officeLiveError } from "../errors.mjs";
import {
  createLiveFailure,
  createLiveSuccess,
  LIVE_PROTOCOL,
  MAX_LIVE_IMAGE_BYTES,
  validateLiveEnvelope,
} from "../protocol.mjs";

const PACKAGE_ROOT = fileURLToPath(new URL("../../..", import.meta.url));
const MAX_ID_LENGTH = 256;
const MAX_TEXT_LENGTH = 100_000;
const MAX_SLIDES_PER_READ = 100;
const SHAPE_TYPES = new Set(["rect", "ellipse", "roundRect", "line", "triangle", "hexagon", "diamond"]);

export const POWERPOINT_BROWSER_COOKIE = "officekit_powerpoint_browser";
export const POWERPOINT_BROWSER_PANE_HEADER = "x-officekit-pane";
export const POWERPOINT_STATIC_PATHS = new Map([
  ["/powerpoint", "taskpane.html"],
  ["/powerpoint/", "taskpane.html"],
  ["/powerpoint/taskpane.html", "taskpane.html"],
  ["/powerpoint/taskpane.js", "taskpane.js"],
  ["/powerpoint/taskpane.css", "taskpane.css"],
  ["/powerpoint/support.html", "support.html"],
  ["/powerpoint/assets/officekit-powerpoint-32.png", "assets/officekit-powerpoint-32.png"],
  ["/powerpoint/assets/officekit-powerpoint-80.png", "assets/officekit-powerpoint-80.png"],
]);

export const POWERPOINT_LIVE_OPERATIONS = Object.freeze([
  "read_presentation",
  "read_slides",
  "read_slide",
  "read_selection",
  "write_text",
  "add_textbox",
  "add_shape",
  "add_image",
  "update_shape",
  "delete_shape",
  "add_slide",
  "read_slide_image",
  "save",
]);

const OPERATION_SET = new Set(POWERPOINT_LIVE_OPERATIONS);

export function createPowerPointLiveAdapter({
  staticRoot = path.join(PACKAGE_ROOT, "apps", "powerpoint-addin", "dist"),
} = {}) {
  return Object.freeze({
    host: "powerpoint",
    sessionIdPrefix: "powerpoint",
    staticRoot,
    staticPaths: POWERPOINT_STATIC_PATHS,
    browserCookie: POWERPOINT_BROWSER_COOKIE,
    browserPaneHeader: POWERPOINT_BROWSER_PANE_HEADER,
    protocol: () => ({
      protocol: LIVE_PROTOCOL,
      operations: POWERPOINT_LIVE_OPERATIONS,
      limits: {
        maxRequestBytes: 10_000_000,
        maxResponseBytes: 10_000_000,
        maxImageBytes: MAX_LIVE_IMAGE_BYTES,
      },
    }),
    error: officeLiveError,
    success: createLiveSuccess,
    failure: createLiveFailure,
    validateRequest: validatePowerPointRequest,
    validateClient: validatePowerPointBrowserClient,
    describeClient: (client) => ({
      presentation: client.presentation,
      capabilities: client.capabilities,
      host: client.host,
    }),
    sameTarget: (left, right) => {
      const leftId = left?.presentation?.id;
      const rightId = right?.presentation?.id;
      return leftId !== undefined && rightId !== undefined
        ? leftId === rightId && left.presentation.name === right.presentation.name
        : left?.presentation?.name === right?.presentation?.name;
    },
    audit: ({ session, record, summary }) => ({
      presentation: session.client.presentation.name,
      operation: record.request.operation,
      target: targetSummary(record.request.args),
      requestHash: summary.requestHash,
      status: record.response?.ok ? "ok" : "error",
    }),
    targetLabel: "presentation",
    unavailableMessage: "PowerPoint session is unavailable. Open OfficeKit in the target presentation and connect it.",
    disconnectedMessage: (reason) => `PowerPoint session disconnected: ${reason}.`,
    routeNotFoundMessage: "PowerPoint bridge route was not found.",
    operationFailureMessage: "PowerPoint operation failed.",
    assetErrorMessage: "PowerPoint add-in assets are missing. Reinstall OfficeKit or run its package build.",
  });
}

export function validatePowerPointBrowserClient(value) {
  if (value == null || typeof value !== "object" || Array.isArray(value)) {
    throw officeLiveError("invalid-session", "PowerPoint add-in client descriptor must be an object.");
  }
  const paneId = identifier(value.paneId, "paneId");
  const presentation = value.presentation;
  if (
    presentation == null || typeof presentation !== "object" || Array.isArray(presentation) ||
    typeof presentation.name !== "string" || presentation.name.length === 0 || presentation.name.length > 512
  ) {
    throw officeLiveError("invalid-session", "PowerPoint add-in did not report a valid presentation target.");
  }
  if (presentation.id !== undefined) identifier(presentation.id, "presentation.id");
  if (presentation.activeSlideId !== undefined) identifier(presentation.activeSlideId, "presentation.activeSlideId");
  if (presentation.slideCount !== undefined && (!Number.isSafeInteger(presentation.slideCount) || presentation.slideCount < 0 || presentation.slideCount > 10_000)) {
    throw officeLiveError("invalid-session", "presentation.slideCount is invalid.");
  }
  return {
    paneId,
    presentation: {
      name: presentation.name,
      ...(presentation.id === undefined ? {} : { id: presentation.id }),
      ...(presentation.activeSlideId === undefined ? {} : { activeSlideId: presentation.activeSlideId }),
      ...(presentation.slideCount === undefined ? {} : { slideCount: presentation.slideCount }),
    },
    capabilities: booleanMap(value.capabilities),
    host: hostMap(value.host),
  };
}

export function validatePowerPointRequest(value) {
  const request = validateLiveEnvelope(value, { host: "powerpoint" });
  if (!OPERATION_SET.has(request.operation)) {
    throw officeLiveError("unsupported-operation", `Unsupported PowerPoint Live operation: ${request.operation}.`);
  }
  validateOperationArguments(request.operation, request.args);
  return request;
}

function validateOperationArguments(operation, args) {
  switch (operation) {
    case "read_presentation":
    case "read_selection":
    case "save":
      assertOnlyKeys(args, [], operation);
      return;
    case "read_slides":
      assertOnlyKeys(args, ["slideIds"], operation);
      if (args.slideIds !== undefined) {
        if (!Array.isArray(args.slideIds) || args.slideIds.length < 1 || args.slideIds.length > MAX_SLIDES_PER_READ) {
          throw invalid(`${operation}.slideIds must contain 1 through ${MAX_SLIDES_PER_READ} slide IDs.`);
        }
        for (const value of args.slideIds) identifier(value, "slideIds entry");
      }
      return;
    case "read_slide":
      assertOnlyKeys(args, ["slideId"], operation);
      identifier(args.slideId, "slideId");
      return;
    case "write_text":
      assertOnlyKeys(args, ["slideId", "shapeId", "text", "expectedText", "expectedSnapshot"], operation);
      target(args);
      text(args.text, "text");
      if (args.expectedText !== undefined) text(args.expectedText, "expectedText");
      expectedSnapshot(args.expectedSnapshot, operation);
      if (args.expectedText !== undefined && args.expectedSnapshot?.text !== undefined) throw invalid(`${operation} accepts expectedText or expectedSnapshot.text, not both.`);
      return;
    case "add_textbox":
      assertOnlyKeys(args, ["slideId", "text", "left", "top", "width", "height"], operation);
      identifier(args.slideId, "slideId");
      text(args.text, "text");
      geometry(args);
      return;
    case "add_shape":
      assertOnlyKeys(args, ["slideId", "type", "left", "top", "width", "height", "text"], operation);
      identifier(args.slideId, "slideId");
      if (!SHAPE_TYPES.has(args.type)) throw invalid("add_shape.type is unsupported.");
      geometry(args);
      if (args.text !== undefined) text(args.text, "text");
      return;
    case "add_image":
      assertOnlyKeys(args, ["slideId", "imageData", "left", "top", "width", "height", "altText"], operation);
      identifier(args.slideId, "slideId");
      assertSafeImageDataUrl(args.imageData);
      geometry(args);
      if (args.altText !== undefined) text(args.altText, "altText", { allowEmpty: true });
      return;
    case "update_shape":
      assertOnlyKeys(args, ["slideId", "shapeId", "left", "top", "width", "height", "text", "expectedText", "expectedSnapshot"], operation);
      target(args);
      if (["left", "top", "width", "height"].some((key) => args[key] !== undefined)) geometry(args, { partial: true });
      if (args.text !== undefined) text(args.text, "text");
      if (args.expectedText !== undefined) text(args.expectedText, "expectedText");
      expectedSnapshot(args.expectedSnapshot, operation);
      if (args.expectedText !== undefined && args.expectedSnapshot?.text !== undefined) throw invalid(`${operation} accepts expectedText or expectedSnapshot.text, not both.`);
      if (["left", "top", "width", "height", "text"].every((key) => args[key] === undefined)) throw invalid("update_shape requires a geometry or text change.");
      return;
    case "delete_shape":
      assertOnlyKeys(args, ["slideId", "shapeId", "expectedSnapshot"], operation);
      target(args);
      expectedSnapshot(args.expectedSnapshot, operation);
      return;
    case "add_slide":
      assertOnlyKeys(args, ["layout", "slideMasterId", "layoutId"], operation);
      if (args.layout !== undefined) text(args.layout, "layout");
      if (args.slideMasterId !== undefined) identifier(args.slideMasterId, "slideMasterId");
      if (args.layoutId !== undefined) identifier(args.layoutId, "layoutId");
      if (args.layout !== undefined && args.layoutId !== undefined) throw invalid("add_slide accepts layout or layoutId, not both.");
      return;
    case "read_slide_image":
      assertOnlyKeys(args, ["slideId", "width", "height"], operation);
      identifier(args.slideId, "slideId");
      positiveInteger(args.width, "width", 4_096);
      positiveInteger(args.height, "height", 4_096);
      return;
    default:
      throw officeLiveError("unsupported-operation", `Unsupported PowerPoint Live operation: ${operation}.`);
  }
}

function target(args) {
  identifier(args.slideId, "slideId");
  identifier(args.shapeId, "shapeId");
}

function geometry(args, { partial = false } = {}) {
  for (const key of ["left", "top", "width", "height"]) {
    if (args[key] === undefined && partial) continue;
    positiveNumber(args[key], key, key === "left" || key === "top");
  }
}

function positiveNumber(value, label, allowZero = false) {
  if (typeof value !== "number" || !Number.isFinite(value) || (allowZero ? value < 0 : value <= 0)) {
    throw invalid(`${label} must be a finite ${allowZero ? "non-negative" : "positive"} number.`);
  }
  if (value > 100_000) throw invalid(`${label} exceeds the live geometry limit.`);
}

function positiveInteger(value, label, maximum) {
  if (!Number.isSafeInteger(value) || value < 1 || value > maximum) throw invalid(`${label} must be an integer from 1 through ${maximum}.`);
}

function identifier(value, label) {
  if (typeof value !== "string" || value.length < 1 || value.length > MAX_ID_LENGTH || !/^[A-Za-z0-9][A-Za-z0-9._:-]*$/u.test(value)) {
    throw invalid(`${label} is invalid.`);
  }
  return value;
}

function text(value, label, { allowEmpty = false } = {}) {
  if (typeof value !== "string" || value.length > MAX_TEXT_LENGTH || (!allowEmpty && value.length === 0)) throw invalid(`${label} must be a bounded string.`);
}

function expectedSnapshot(value, operation) {
  if (value === undefined) return;
  if (value == null || typeof value !== "object" || Array.isArray(value)) throw invalid(`${operation}.expectedSnapshot must be an object.`);
  assertOnlyKeys(value, ["id", "name", "type", "text", "left", "top", "width", "height"], `${operation}.expectedSnapshot`);
  if (value.id !== undefined) identifier(value.id, `${operation}.expectedSnapshot.id`);
  if (value.name !== undefined) text(value.name, `${operation}.expectedSnapshot.name`, { allowEmpty: true });
  if (value.type !== undefined) text(value.type, `${operation}.expectedSnapshot.type`);
  if (value.text !== undefined) text(value.text, `${operation}.expectedSnapshot.text`, { allowEmpty: true });
  for (const key of ["left", "top", "width", "height"]) {
    if (value[key] !== undefined) positiveNumber(value[key], `${operation}.expectedSnapshot.${key}`, key === "left" || key === "top");
  }
}

function assertSafeImageDataUrl(value) {
  if (typeof value !== "string" || !/^data:image\/(png|jpeg|gif|svg\+xml);base64,[A-Za-z0-9+/=]+$/u.test(value) || Buffer.byteLength(value, "utf8") > MAX_LIVE_IMAGE_BYTES) {
    throw invalid("add_image.imageData must be a bounded base64 PNG, JPEG, GIF, or safe SVG data URL.");
  }
  if (!value.startsWith("data:image/svg+xml;")) return;
  let svg;
  try {
    svg = Buffer.from(value.slice(value.indexOf(",") + 1), "base64").toString("utf8");
  } catch {
    throw invalid("add_image.imageData contains an invalid SVG payload.");
  }
  if (/<\s*script\b|\bon[a-z]+\s*=|(?:href|xlink:href)\s*=\s*["']\s*(?:https?:|file:|data:)/iu.test(svg)) {
    throw invalid("add_image.imageData contains an unsafe SVG reference or script.");
  }
}

function assertOnlyKeys(value, keys, operation) {
  const allowed = new Set(keys);
  for (const key of Object.keys(value)) if (!allowed.has(key)) throw invalid(`${operation}.${key} is not supported.`);
}

function invalid(message) {
  return officeLiveError("invalid-request", message);
}

function booleanMap(value) {
  return value != null && typeof value === "object" && !Array.isArray(value)
    ? Object.fromEntries(Object.entries(value).filter(([, candidate]) => typeof candidate === "boolean"))
    : {};
}

function hostMap(value) {
  return value != null && typeof value === "object" && !Array.isArray(value)
    ? {
      platform: typeof value.platform === "string" ? value.platform.slice(0, 128) : "unknown",
      version: typeof value.version === "string" ? value.version.slice(0, 128) : "unknown",
      webView: typeof value.webView === "string" ? value.webView.slice(0, 128) : "unknown",
    }
    : { platform: "unknown", version: "unknown", webView: "unknown" };
}

function targetSummary(args) {
  if (args?.slideId && args?.shapeId) return `${args.slideId}/${args.shapeId}`;
  if (args?.slideId) return String(args.slideId);
  return undefined;
}
