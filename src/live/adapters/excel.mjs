import path from "node:path";
import { fileURLToPath } from "node:url";

import { excelLiveError } from "../../excel-live/errors.mjs";
import {
  createExcelFailure,
  createExcelSuccess,
  protocolReference,
  validateExcelRequest,
} from "../../excel-live/protocol.mjs";

const PACKAGE_ROOT = fileURLToPath(new URL("../../..", import.meta.url));

export const EXCEL_BROWSER_COOKIE = "officekit_excel_browser";
export const EXCEL_BROWSER_PANE_HEADER = "x-officekit-pane";
export const EXCEL_STATIC_PATHS = new Map([
  ["/", "taskpane.html"],
  ["/taskpane.html", "taskpane.html"],
  ["/taskpane.js", "taskpane.js"],
  ["/taskpane.css", "taskpane.css"],
  ["/support.html", "support.html"],
  ["/assets/officekit-excel-32.png", "assets/officekit-excel-32.png"],
  ["/assets/officekit-excel-80.png", "assets/officekit-excel-80.png"],
]);

export function createExcelLiveAdapter({ staticRoot = path.join(PACKAGE_ROOT, "apps", "excel-addin", "dist") } = {}) {
  return Object.freeze({
    host: "excel",
    staticRoot,
    staticPaths: EXCEL_STATIC_PATHS,
    browserCookie: EXCEL_BROWSER_COOKIE,
    browserPaneHeader: EXCEL_BROWSER_PANE_HEADER,
    protocol: protocolReference,
    error: excelLiveError,
    success: createExcelSuccess,
    failure: createExcelFailure,
    validateRequest: validateExcelLiveRequest,
    validateClient: validateExcelBrowserClient,
    describeClient: (client) => ({
      workbook: client.workbook,
      capabilities: client.capabilities,
      host: client.host,
    }),
    sameTarget: (left, right) => left?.workbook?.name === right?.workbook?.name,
    audit: ({ session, record, summary }) => ({
      workbook: session.client.workbook.name,
      operation: record.request.operation,
      range: rangeSummary(record.request.args),
      requestHash: summary.requestHash,
      status: record.response?.ok ? "ok" : "error",
    }),
    targetLabel: "workbook",
    unavailableMessage: "Excel session is unavailable. Open OfficeKit in the target workbook and connect it.",
    disconnectedMessage: (reason) => `Excel session disconnected: ${reason}.`,
    routeNotFoundMessage: "Excel bridge route was not found.",
    operationFailureMessage: "Excel operation failed.",
    assetErrorMessage: "Excel add-in assets are missing. Reinstall OfficeKit or run its package build.",
  });
}

function validateExcelLiveRequest(value) {
  if (value?.host !== undefined && value.host !== "excel") {
    throw excelLiveError("forbidden", "Live request is not for the Excel session.");
  }
  return validateExcelRequest(value);
}

export function validateExcelBrowserClient(value) {
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
  return {
    paneId,
    workbook: { name: workbook.name, activeSheet: workbook.activeSheet },
    capabilities: booleanMap(value.capabilities),
    host: stringMap(value.host),
  };
}

function booleanMap(value) {
  return value != null && typeof value === "object" && !Array.isArray(value)
    ? Object.fromEntries(Object.entries(value).filter(([, candidate]) => typeof candidate === "boolean"))
    : {};
}

function stringMap(value) {
  return value != null && typeof value === "object" && !Array.isArray(value)
    ? {
      platform: typeof value.platform === "string" ? value.platform.slice(0, 128) : "unknown",
      version: typeof value.version === "string" ? value.version.slice(0, 128) : "unknown",
      webView: typeof value.webView === "string" ? value.webView.slice(0, 128) : "unknown",
    }
    : { platform: "unknown", version: "unknown", webView: "unknown" };
}

function rangeSummary(args) {
  const values = [];
  for (const candidate of [args, args?.source, args?.destination]) {
    if (candidate?.sheet && candidate?.range) values.push(`${candidate.sheet}!${candidate.range}`);
  }
  return values.join(",") || undefined;
}
