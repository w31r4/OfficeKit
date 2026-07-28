/* global Office, Excel */

declare const Office: any;
declare const Excel: any;

type Json = null | boolean | number | string | Json[] | { [key: string]: Json };
type Session = { id: string; workbook: { name: string; activeSheet: string }; capabilities: Record<string, boolean> };
type OperationEnvelope = { requestId: string; request: { operation: string; args: Record<string, any> } };

const paneId = taskPaneIdentity();
let activeSession: Session | null = null;
let polling = false;
let maxImageBytes = 8_000_000;
let maxResultBytes = 9_000_000;
const maxReadCells = 20_000;
const maxSearchCells = 25_000;
const maxImageRangeCells = 2_500;

const connectButton = byId<HTMLButtonElement>("connect");
const disconnectButton = byId<HTMLButtonElement>("disconnect");
const refreshButton = byId<HTMLButtonElement>("refresh");
const setup = byId<HTMLElement>("setup");
const setupCopy = byId<HTMLElement>("setup-copy");
const sessionCard = byId<HTMLElement>("session");
const workbookName = byId<HTMLElement>("workbook-name");
const sessionId = byId<HTMLElement>("session-id");
const connectionStatus = byId<HTMLElement>("connection-status");
const diagnosticsOutput = byId<HTMLElement>("diagnostics-output");
const auditOutput = byId<HTMLElement>("audit-output");

Office.onReady(async (info: { host?: string }) => {
  if (info.host !== Office.HostType.Excel) {
    setSetup("Open OfficeKit from Microsoft Excel.", false);
    return;
  }
  try {
    const bootstrap = await requestJson("/v1/browser/bootstrap");
    maxImageBytes = bootstrap?.limits?.maxImageBytes ?? maxImageBytes;
    maxResultBytes = (bootstrap?.limits?.maxImageBytes ?? maxImageBytes) + (bootstrap?.limits?.maxRequestBytes ?? 1_000_000);
    showDiagnostics(bootstrap);
    connectButton.disabled = false;
    setSetup("Click Connect OfficeKit to make this open workbook available to your Agent.", true);
  } catch (error) {
    setSetup(`OfficeKit bridge is unavailable: ${messageOf(error)}`, false);
    showDiagnostics({ bridge: "unavailable", error: messageOf(error) });
  }
});

connectButton.addEventListener("click", () => void connect());
disconnectButton.addEventListener("click", () => void disconnect());
refreshButton.addEventListener("click", () => void refresh());

async function connect(): Promise<void> {
  connectButton.disabled = true;
  setSetup("Reading this workbook…", false);
  try {
    const client = await workbookDescriptor();
    const response = await requestJson("/v1/browser/sessions", { client });
    activeSession = response.session as Session;
    showSession(activeSession);
    void poll(activeSession.id);
  } catch (error) {
    setSetup(`Could not connect: ${messageOf(error)}`, true);
  } finally {
    connectButton.disabled = false;
  }
}

async function refresh(): Promise<void> {
  if (activeSession == null) return;
  try {
    const client = await workbookDescriptor();
    const response = await requestJson(
      `/v1/browser/sessions/${encodeURIComponent(activeSession.id)}/refresh`,
      { client },
    );
    activeSession = response.session as Session;
    showSession(activeSession);
  } catch (error) {
    connectionStatus.textContent = `Refresh failed: ${messageOf(error)}`;
  }
}

async function disconnect(): Promise<void> {
  if (activeSession == null) return;
  try {
    await requestJson(`/v1/browser/sessions/${encodeURIComponent(activeSession.id)}/disconnect`, {}, "POST");
  } catch {
    // The server may already be gone. Locally forgetting the session is still correct.
  }
  activeSession = null;
  polling = false;
  sessionCard.hidden = true;
  setup.hidden = false;
  setSetup("Disconnected. Click Connect OfficeKit to share this workbook again.", true);
}

async function poll(id: string): Promise<void> {
  if (polling) return;
  polling = true;
  while (activeSession?.id === id) {
    try {
      const response = await fetch(`/v1/browser/sessions/${encodeURIComponent(id)}/next`, {
        credentials: "same-origin",
        cache: "no-store",
        headers: { "x-officekit-pane": paneId },
      });
      if (response.status === 204) continue;
      if (response.status === 404 || response.status === 410) {
        await disconnect();
        break;
      }
      const body = await parseResponse(response) as OperationEnvelope;
      connectionStatus.textContent = `Running ${body.request.operation}…`;
      const result = await executeOperation(body.request);
      const accepted = await submitResult(id, body.requestId, result);
      showAudit(accepted?.completion);
      connectionStatus.textContent = "Waiting for OfficeKit commands";
    } catch (error) {
      connectionStatus.textContent = `Reconnecting: ${messageOf(error)}`;
      await delay(900);
    }
  }
  polling = false;
}

async function workbookDescriptor(): Promise<Record<string, Json>> {
  return Excel.run(async (context: any) => {
    const workbook = context.workbook;
    const active = workbook.worksheets.getActiveWorksheet();
    workbook.load("name");
    active.load("name");
    await context.sync();
    return {
      workbook: {
        name: typeof workbook.name === "string" && workbook.name ? workbook.name : "Unsaved workbook",
        activeSheet: active.name,
      },
      paneId,
      host: {
        platform: Office.context.platform,
        version: Office.context.diagnostics?.version ?? "unknown",
        webView: detectedWebView(),
      },
      capabilities: {
        excelApi18: Office.context.requirements.isSetSupported("ExcelApi", "1.8"),
        sharedRuntime: Office.context.requirements.isSetSupported("SharedRuntime", "1.1"),
        rangeImage: Office.context.requirements.isSetSupported("ExcelApi", "1.8"),
        save: Office.context.requirements.isSetSupported("ExcelApi", "1.11"),
        desktopWindow: Office.context.requirements.isSetSupported("ExcelApiDesktop", "1.1"),
      },
    };
  });
}

async function executeOperation(request: { operation: string; args: Record<string, any> }): Promise<{ ok: boolean; result?: Json; error?: Record<string, Json> }> {
  try {
    requireOperationCapability(request.operation, request.args);
    const result = await Excel.run(async (context: any) => executeWithContext(context, request.operation, request.args));
    if (jsonByteLength(result) > maxResultBytes) {
      const error: any = new Error(`Excel operation result exceeds the ${maxResultBytes}-byte bridge limit.`);
      error.code = "response-too-large";
      throw error;
    }
    return { ok: true, result: result ?? {} };
  } catch (error: any) {
    const code = ["unsupported-capability", "response-too-large"].includes(error?.code)
      ? error.code
      : /ApiNotFound|NotImplemented|Unsupported/i.test(String(error?.code ?? error?.message ?? ""))
        ? "unsupported-capability"
        : "office-operation-failed";
    return {
      ok: false,
      error: {
        code,
        message: messageOf(error),
        retryable: code !== "unsupported-capability",
      },
    };
  }
}

function requireOperationCapability(operation: string, args: Record<string, any>): void {
  const requirement = OPERATION_REQUIREMENTS[operation] ?? { set: "ExcelApi", version: "1.8" };
  if (!Office.context.requirements.isSetSupported(requirement.set, requirement.version)) {
    const error: any = new Error(`${operation} requires ${requirement.set} ${requirement.version} or later in this Excel client.`);
    error.code = "unsupported-capability";
    throw error;
  }
  if (operation === "update_sheet_view" && args.zoom !== undefined && !Office.context.requirements.isSetSupported("ExcelApiDesktop", "1.1")) {
    const error: any = new Error("update_sheet_view.zoom requires ExcelApiDesktop 1.1 or later in this Excel client.");
    error.code = "unsupported-capability";
    throw error;
  }
}

const OPERATION_REQUIREMENTS: Record<string, { set: string; version: string }> = {
  save: { set: "ExcelApi", version: "1.11" },
  read_ranges: { set: "ExcelApi", version: "1.8" },
  search_workbook: { set: "ExcelApi", version: "1.8" },
  list_items: { set: "ExcelApi", version: "1.8" },
  write_range: { set: "ExcelApi", version: "1.8" },
  clear_range: { set: "ExcelApi", version: "1.8" },
  update_sheet: { set: "ExcelApi", version: "1.8" },
  update_workbook: { set: "ExcelApi", version: "1.8" },
  copy_range_to: { set: "ExcelApi", version: "1.8" },
  read_range_image: { set: "ExcelApi", version: "1.8" },
  read_sheets_metadata: { set: "ExcelApi", version: "1.8" },
  resize_range: { set: "ExcelApi", version: "1.8" },
  update_sheet_view: { set: "ExcelApi", version: "1.8" },
  format_range: { set: "ExcelApi", version: "1.8" },
  chart: { set: "ExcelApi", version: "1.8" },
  table: { set: "ExcelApi", version: "1.8" },
  pivot_table: { set: "ExcelApi", version: "1.8" },
};

async function executeWithContext(context: any, operation: string, args: Record<string, any>): Promise<Json> {
  switch (operation) {
    case "read_ranges": return readRanges(context, args);
    case "search_workbook": return searchWorkbook(context, args);
    case "list_items": return listItems(context, args);
    case "write_range": return writeRange(context, args);
    case "clear_range": return clearRange(context, args);
    case "update_sheet": return updateSheet(context, args);
    case "update_workbook": return updateWorkbook(context, args);
    case "copy_range_to": return copyRangeTo(context, args);
    case "read_range_image": return readRangeImage(context, args);
    case "read_sheets_metadata": return readSheetsMetadata(context);
    case "resize_range": return resizeRange(context, args);
    case "update_sheet_view": return updateSheetView(context, args);
    case "format_range": return formatRange(context, args);
    case "chart": return chart(context, args);
    case "table": return table(context, args);
    case "pivot_table": return pivotTable(context, args);
    case "save":
      // An unsaved workbook must use Excel's own Save As prompt. OfficeKit
      // never selects a path or silently commits a new file on the user's behalf.
      context.workbook.save(Excel.SaveBehavior.prompt);
      await context.sync();
      return { saveRequested: true };
    default:
      throw new Error(`Unsupported OfficeKit Excel operation: ${operation}`);
  }
}

async function readRanges(context: any, args: any): Promise<Json> {
  const sheet = context.workbook.worksheets.getItem(args.sheet);
  const ranges = args.ranges.map((address: string) => sheet.getRange(address));
  const include = args.include ?? ["values", "formulas", "text", "numberFormat"];
  for (const range of ranges) range.load("rowCount,columnCount");
  await context.sync();
  const cells = ranges.reduce((total: number, range: any) => total + range.rowCount * range.columnCount, 0);
  if (cells > maxReadCells) throw boundedOperationError("read_ranges", maxReadCells, "cells");
  for (const range of ranges) range.load(include.join(","));
  await context.sync();
  return {
    sheet: args.sheet,
    ranges: ranges.map((range: any, index: number) => ({
      address: args.ranges[index],
      ...(include.includes("values") ? { values: range.values } : {}),
      ...(include.includes("formulas") ? { formulas: range.formulas } : {}),
      ...(include.includes("text") ? { text: range.text } : {}),
      ...(include.includes("numberFormat") ? { numberFormat: range.numberFormat } : {}),
    })),
  };
}

async function searchWorkbook(context: any, args: any): Promise<Json> {
  const options = args.options ?? {};
  const maximum = Math.min(Number.isSafeInteger(options.maxResults) ? options.maxResults : 100, 500);
  const worksheets = context.workbook.worksheets;
  worksheets.load("items/name");
  await context.sync();
  const ranges = worksheets.items.map((sheet: any) => sheet.getUsedRangeOrNullObject());
  for (const range of ranges) range.load("isNullObject,rowCount,columnCount,text,address");
  await context.sync();
  let cells = 0;
  for (const range of ranges) {
    if (!range.isNullObject) cells += range.rowCount * range.columnCount;
  }
  if (cells > maxSearchCells) throw boundedOperationError("search_workbook", maxSearchCells, "cells");
  for (const range of ranges) if (!range.isNullObject) range.load("text");
  await context.sync();
  const matcher = makeMatcher(args.query, options);
  const matches: Json[] = [];
  for (let worksheetIndex = 0; worksheetIndex < worksheets.items.length; worksheetIndex += 1) {
    const range = ranges[worksheetIndex];
    if (range.isNullObject) continue;
    for (let row = 0; row < range.text.length && matches.length < maximum; row += 1) {
      for (let column = 0; column < range.text[row].length && matches.length < maximum; column += 1) {
        const text = String(range.text[row][column] ?? "");
        if (matcher(text)) {
          matches.push({ sheet: worksheets.items[worksheetIndex].name, address: cellAddress(range.address, row, column), text });
        }
      }
    }
  }
  return { query: args.query, matches, truncated: matches.length >= maximum };
}

async function listItems(context: any, args: any): Promise<Json> {
  if (args.kind === "worksheets") return readSheetsMetadata(context);
  if (args.kind === "tables") {
    const tables = context.workbook.tables;
    tables.load("items/name,items/id,items/showHeaders,items/showTotals");
    await context.sync();
    return { kind: args.kind, items: tables.items.map((item: any) => ({ name: item.name, id: item.id, showHeaders: item.showHeaders, showTotals: item.showTotals })) };
  }
  if (args.kind === "pivotTables") {
    const pivots = context.workbook.pivotTables;
    pivots.load("items/name,items/id");
    await context.sync();
    return { kind: args.kind, items: pivots.items.map((item: any) => ({ name: item.name, id: item.id })) };
  }
  if (args.kind === "names") {
    const names = context.workbook.names;
    names.load("items/name,items/formula,items/value");
    await context.sync();
    return { kind: args.kind, items: names.items.map((item: any) => ({ name: item.name, formula: item.formula, value: item.value })) };
  }
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();
  const collections = sheets.items.map((sheet: any) => sheet.charts);
  for (const collection of collections) collection.load("items/name,items/id");
  await context.sync();
  return {
    kind: args.kind,
    items: collections.flatMap((collection: any, index: number) => collection.items.map((item: any) => ({ sheet: sheets.items[index].name, name: item.name, id: item.id }))),
  };
}

async function writeRange(context: any, args: any): Promise<Json> {
  const range = sheetRange(context, args);
  if (args.values !== undefined) range.values = args.values;
  if (args.formulas !== undefined) range.formulas = args.formulas;
  if (args.numberFormat !== undefined) range.numberFormat = args.numberFormat;
  await context.sync();
  return { sheet: args.sheet, range: args.range, written: true };
}

async function clearRange(context: any, args: any): Promise<Json> {
  const applies = { all: Excel.ClearApplyTo.all, contents: Excel.ClearApplyTo.contents, formats: Excel.ClearApplyTo.formats, hyperlinks: Excel.ClearApplyTo.hyperlinks, removeHyperlinks: Excel.ClearApplyTo.removeHyperlinks };
  const applyTo = (args.applyTo ?? "all") as keyof typeof applies;
  sheetRange(context, args).clear(applies[applyTo]);
  await context.sync();
  return { sheet: args.sheet, range: args.range, cleared: args.applyTo ?? "all" };
}

async function updateSheet(context: any, args: any): Promise<Json> {
  const sheets = context.workbook.worksheets;
  if (args.action === "add") sheets.add(args.name);
  else {
    const sheet = sheets.getItem(args.name);
    if (args.action === "rename") sheet.name = args.newName;
    if (args.action === "delete") sheet.delete();
    if (args.action === "activate") sheet.activate();
  }
  await context.sync();
  return { action: args.action, name: args.action === "rename" ? args.newName : args.name };
}

async function updateWorkbook(context: any, args: any): Promise<Json> {
  context.application.calculationMode = args.calculationMode;
  await context.sync();
  return { calculationMode: args.calculationMode };
}

async function copyRangeTo(context: any, args: any): Promise<Json> {
  const source = sheetRange(context, args.source);
  const destination = sheetRange(context, args.destination);
  const copyType = { all: Excel.RangeCopyType.all, values: Excel.RangeCopyType.values, formulas: Excel.RangeCopyType.formulas, formats: Excel.RangeCopyType.formats };
  const type = (args.type ?? "all") as keyof typeof copyType;
  destination.copyFrom(source, copyType[type]);
  await context.sync();
  return { copied: true, source: args.source, destination: args.destination };
}

async function readRangeImage(context: any, args: any): Promise<Json> {
  const range = sheetRange(context, args);
  range.load("rowCount,columnCount");
  await context.sync();
  if (range.rowCount * range.columnCount > maxImageRangeCells) {
    throw boundedOperationError("read_range_image", maxImageRangeCells, "cells");
  }
  const image = range.getImage();
  await context.sync();
  if (image.value.length > Math.ceil(maxImageBytes * 4 / 3)) {
    throw new Error(`read_range_image exceeds the ${maxImageBytes}-byte image limit.`);
  }
  return { sheet: args.sheet, range: args.range, dataUri: `data:image/png;base64,${image.value}` };
}

async function readSheetsMetadata(context: any): Promise<Json> {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name,items/id,items/position,items/visibility");
  await context.sync();
  return { sheets: sheets.items.map((sheet: any) => ({ name: sheet.name, id: sheet.id, position: sheet.position, visibility: sheet.visibility })) };
}

async function resizeRange(context: any, args: any): Promise<Json> {
  const range = sheetRange(context, args);
  if (args.columnWidth !== undefined) range.format.columnWidth = args.columnWidth;
  if (args.rowHeight !== undefined) range.format.rowHeight = args.rowHeight;
  if (args.autofitColumns) range.format.autofitColumns();
  if (args.autofitRows) range.format.autofitRows();
  await context.sync();
  return { sheet: args.sheet, range: args.range, resized: true };
}

async function updateSheetView(context: any, args: any): Promise<Json> {
  const sheet = context.workbook.worksheets.getItem(args.sheet);
  if (args.freezeRows !== undefined || args.freezeColumns !== undefined) {
    sheet.freezePanes.unfreeze();
    if (args.freezeRows) sheet.freezePanes.freezeRows(args.freezeRows);
    if (args.freezeColumns) sheet.freezePanes.freezeColumns(args.freezeColumns);
  }
  if (args.showGridlines !== undefined) sheet.showGridlines = args.showGridlines;
  if (args.zoom !== undefined) context.workbook.application.activeWindow.zoom = args.zoom;
  if (args.selectRange !== undefined) {
    sheet.activate();
    sheet.getRange(args.selectRange).select();
  }
  await context.sync();
  return { sheet: args.sheet, updated: true };
}

async function formatRange(context: any, args: any): Promise<Json> {
  const range = sheetRange(context, args);
  const format = args.format;
  if (format.numberFormat !== undefined) range.numberFormat = format.numberFormat;
  if (format.fill?.color !== undefined) range.format.fill.color = format.fill.color;
  if (format.font !== undefined) Object.assign(range.format.font, format.font);
  for (const property of ["horizontalAlignment", "verticalAlignment", "wrapText", "columnWidth", "rowHeight", "indentLevel", "textOrientation"]) {
    if (format[property] !== undefined) range.format[property] = format[property];
  }
  if (Array.isArray(format.borders)) {
    for (const border of format.borders) {
      const target = range.format.borders.getItem(border.side);
      if (border.style !== undefined) target.style = border.style;
      if (border.color !== undefined) target.color = border.color;
    }
  }
  await context.sync();
  return { sheet: args.sheet, range: args.range, formatted: true };
}

async function chart(context: any, args: any): Promise<Json> {
  const sheet = context.workbook.worksheets.getItem(args.sheet);
  if (args.action === "create") {
    const created = sheet.charts.add(args.type, sheet.getRange(args.sourceRange), Excel.ChartSeriesBy.auto);
    created.name = args.name;
    if (args.title !== undefined) created.title.text = args.title;
    if (args.position?.start && args.position?.end) created.setPosition(args.position.start, args.position.end);
  } else {
    const existing = sheet.charts.getItem(args.name);
    if (args.action === "delete") existing.delete();
    else {
      if (args.title !== undefined) existing.title.text = args.title;
      if (args.sourceRange !== undefined) existing.setData(sheet.getRange(args.sourceRange), Excel.ChartSeriesBy.auto);
      if (args.position?.start && args.position?.end) existing.setPosition(args.position.start, args.position.end);
    }
  }
  await context.sync();
  return { action: args.action, sheet: args.sheet, name: args.name };
}

async function table(context: any, args: any): Promise<Json> {
  const sheet = context.workbook.worksheets.getItem(args.sheet);
  if (args.action === "create") {
    const created = sheet.tables.add(args.range, args.hasHeaders ?? true);
    created.name = args.name;
  } else {
    const existing = sheet.tables.getItem(args.name);
    if (args.action === "delete") existing.delete();
    if (args.action === "add_rows") existing.rows.add(null, args.rows);
    if (args.action === "delete_rows") {
      for (let index = 0; index < args.count; index += 1) existing.rows.getItemAt(args.index).delete();
    }
  }
  await context.sync();
  return { action: args.action, sheet: args.sheet, name: args.name };
}

async function pivotTable(context: any, args: any): Promise<Json> {
  const pivots = context.workbook.pivotTables;
  if (args.action === "create") pivots.add(args.name, args.source, `${args.sheet}!${args.destination}`);
  else {
    const existing = pivots.getItem(args.name);
    if (args.action === "delete") existing.delete();
    if (args.action === "refresh") existing.refresh();
  }
  await context.sync();
  return { action: args.action, sheet: args.sheet, name: args.name };
}

function sheetRange(context: any, args: any): any {
  return context.workbook.worksheets.getItem(args.sheet).getRange(args.range);
}

function makeMatcher(query: string, options: any): (value: string) => boolean {
  const wanted = options.matchCase ? query : query.toLocaleLowerCase();
  return (value) => {
    const candidate = options.matchCase ? value : value.toLocaleLowerCase();
    return options.completeMatch ? candidate === wanted : candidate.includes(wanted);
  };
}

async function submitResult(
  sessionId: string,
  requestId: string,
  result: { ok: boolean; result?: Json; error?: Record<string, Json> },
): Promise<any> {
  let payload: Record<string, Json> = { requestId, ...result };
  if (jsonByteLength(payload) > maxResultBytes) {
    payload = {
      requestId,
      ok: false,
      error: {
        code: "response-too-large",
        message: `Excel operation result exceeds the ${maxResultBytes}-byte bridge limit.`,
        retryable: false,
      },
    };
  }
  return requestJson(`/v1/browser/sessions/${encodeURIComponent(sessionId)}/results`, payload);
}

function boundedOperationError(operation: string, limit: number, unit: string): Error {
  const error: any = new Error(`${operation} exceeds the ${limit}-${unit} live operation limit.`);
  error.code = "response-too-large";
  return error;
}

function jsonByteLength(value: unknown): number {
  return new TextEncoder().encode(JSON.stringify(value)).byteLength;
}

function detectedWebView(): string {
  const userAgent = typeof navigator === "undefined" ? "" : navigator.userAgent;
  if (/WebView2|Edg\//iu.test(userAgent)) return "webview2-or-edge";
  if (/AppleWebKit/iu.test(userAgent)) return "webkit";
  return "unknown";
}

function cellAddress(usedRangeAddress: string, rowOffset: number, columnOffset: number): string {
  const match = /!\$?([A-Z]+)\$?(\d+)(?::\$?[A-Z]+\$?\d+)?$/u.exec(usedRangeAddress);
  if (match == null) return usedRangeAddress;
  return `${columnName(columnNumber(match[1]) + columnOffset)}${Number(match[2]) + rowOffset}`;
}

function columnNumber(name: string): number {
  return [...name].reduce((number, character) => number * 26 + character.charCodeAt(0) - 64, 0);
}

function columnName(number: number): string {
  let current = number;
  let result = "";
  while (current > 0) {
    current -= 1;
    result = String.fromCharCode(65 + (current % 26)) + result;
    current = Math.floor(current / 26);
  }
  return result;
}

async function requestJson(url: string, body?: Json, method = "POST"): Promise<any> {
  const response = await fetch(url, {
    method,
    credentials: "same-origin",
    cache: "no-store",
    headers: {
      "x-officekit-pane": paneId,
      ...(body === undefined ? {} : { "content-type": "application/json" }),
    },
    ...(body === undefined ? {} : { body: JSON.stringify(body) }),
  });
  return parseResponse(response);
}

async function parseResponse(response: Response): Promise<any> {
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(payload?.error?.message ?? `OfficeKit bridge returned ${response.status}.`);
  return payload;
}

function setSetup(message: string, canConnect: boolean): void {
  setupCopy.textContent = message;
  setup.hidden = false;
  sessionCard.hidden = true;
  connectButton.disabled = !canConnect;
}

function showSession(session: Session): void {
  setup.hidden = true;
  sessionCard.hidden = false;
  workbookName.textContent = `${session.workbook.name} — ${session.workbook.activeSheet}`;
  sessionId.textContent = session.id;
  connectionStatus.textContent = "Waiting for OfficeKit commands";
  showDiagnostics({ session, bridge: "connected" });
}

function showDiagnostics(value: unknown): void {
  diagnosticsOutput.textContent = JSON.stringify(value, null, 2);
}

function showAudit(value: unknown): void {
  auditOutput.textContent = value == null
    ? "No completed request in this task pane."
    : JSON.stringify(value, null, 2);
}

function taskPaneIdentity(): string {
  return typeof crypto?.randomUUID === "function"
    ? crypto.randomUUID()
    : `pane-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function byId<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (element == null) throw new Error(`Missing OfficeKit element: ${id}`);
  return element as T;
}

function delay(milliseconds: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
