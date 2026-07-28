import { excelLiveError } from "./errors.mjs";

export const EXCEL_LIVE_PROTOCOL = 1;
export const MAX_REQUEST_BYTES = 1_000_000;
export const MAX_RANGE_COUNT = 32;
export const MAX_MATRIX_CELLS = 50_000;
export const MAX_IMAGE_BYTES = 8_000_000;
export const MAX_SEARCH_RESULTS = 500;

export const EXCEL_LIVE_OPERATIONS = Object.freeze([
  "read_ranges",
  "search_workbook",
  "list_items",
  "write_range",
  "clear_range",
  "update_sheet",
  "update_workbook",
  "copy_range_to",
  "read_range_image",
  "read_sheets_metadata",
  "resize_range",
  "update_sheet_view",
  "format_range",
  "chart",
  "table",
  "pivot_table",
  "save",
]);

const OPERATION_SET = new Set(EXCEL_LIVE_OPERATIONS);
const MUTATING_OPERATION_SET = new Set([
  "write_range",
  "clear_range",
  "update_sheet",
  "update_workbook",
  "copy_range_to",
  "resize_range",
  "update_sheet_view",
  "format_range",
  "chart",
  "table",
  "pivot_table",
  "save",
]);
const ITEM_KINDS = new Set(["worksheets", "tables", "charts", "pivotTables", "names"]);
const CLEAR_APPLIES_TO = new Set(["all", "contents", "formats", "hyperlinks", "removeHyperlinks"]);
const SHEET_ACTIONS = new Set(["add", "rename", "delete", "activate"]);
const CHART_ACTIONS = new Set(["create", "update", "delete"]);
const TABLE_ACTIONS = new Set(["create", "delete", "add_rows", "delete_rows"]);
const PIVOT_ACTIONS = new Set(["create", "delete", "refresh"]);

export function isMutatingExcelOperation(operation) {
  return MUTATING_OPERATION_SET.has(operation);
}

export function validateExcelRequest(value) {
  assertPlainObject(value, "Excel request");
  assertExactNumber(value.protocol, EXCEL_LIVE_PROTOCOL, "protocol");
  assertIdentifier(value.sessionId, "sessionId", 128);
  assertIdentifier(value.idempotencyKey, "idempotencyKey", 160);
  if (!OPERATION_SET.has(value.operation)) {
    throw excelLiveError(
      "unsupported-operation",
      `Unsupported Excel Live operation: ${String(value.operation)}.`,
    );
  }
  assertPlainObject(value.args, "args");
  const request = cloneJson(value, "Excel request");
  validateOperationArguments(request.operation, request.args);
  return request;
}

export function createExcelSuccess({ result, audit }) {
  return {
    protocol: EXCEL_LIVE_PROTOCOL,
    ok: true,
    result: result ?? {},
    audit: audit ?? {},
  };
}

export function createExcelFailure(error, { audit } = {}) {
  const normalized = error?.code
    ? error
    : excelLiveError(
      "internal-error",
      error instanceof Error ? error.message : String(error),
    );
  return {
    protocol: EXCEL_LIVE_PROTOCOL,
    ok: false,
    error: {
      code: normalized.code,
      message: normalized.message,
      retryable: Boolean(normalized.retryable),
      maybeApplied: Boolean(normalized.maybeApplied),
      ...(normalized.details === undefined ? {} : { details: normalized.details }),
    },
    ...(audit == null ? {} : { audit }),
  };
}

export function protocolReference() {
  return {
    protocol: EXCEL_LIVE_PROTOCOL,
    operations: EXCEL_LIVE_OPERATIONS,
    limits: {
      maxRequestBytes: MAX_REQUEST_BYTES,
      maxRangeCount: MAX_RANGE_COUNT,
      maxMatrixCells: MAX_MATRIX_CELLS,
      maxImageBytes: MAX_IMAGE_BYTES,
    },
  };
}

function validateOperationArguments(operation, args) {
  switch (operation) {
    case "read_ranges":
      assertOnlyKeys(args, ["sheet", "ranges", "include"], operation);
      assertSheet(args);
      assertArray(args.ranges, "ranges", 1, MAX_RANGE_COUNT);
      for (const range of args.ranges) assertA1Range(range, "ranges entry");
      if (args.include !== undefined) {
        assertArray(args.include, "include", 1, 4);
        for (const property of args.include) {
          if (!["values", "formulas", "text", "numberFormat"].includes(property)) {
            throw invalid("include must contain values, formulas, text, or numberFormat.");
          }
        }
      }
      return;
    case "search_workbook":
      assertOnlyKeys(args, ["query", "options"], operation);
      assertString(args.query, "query", 1, 512);
      if (args.options !== undefined) validateSearchOptions(args.options);
      return;
    case "list_items":
      assertOnlyKeys(args, ["kind"], operation);
      if (!ITEM_KINDS.has(args.kind)) throw invalid("list_items.kind is invalid.");
      return;
    case "write_range":
      assertOnlyKeys(args, ["sheet", "range", "values", "formulas", "numberFormat"], operation);
      assertRangeTarget(args);
      if (args.values === undefined && args.formulas === undefined) {
        throw invalid("write_range requires values or formulas.");
      }
      if (args.values !== undefined) assertMatrix(args.values, "values");
      if (args.formulas !== undefined) assertMatrix(args.formulas, "formulas", { stringsOnly: true });
      if (args.numberFormat !== undefined) assertMatrix(args.numberFormat, "numberFormat", { stringsOnly: true });
      assertMatchingMatrices(args.values, args.formulas, args.numberFormat);
      return;
    case "clear_range":
      assertOnlyKeys(args, ["sheet", "range", "applyTo"], operation);
      assertRangeTarget(args);
      if (args.applyTo !== undefined && !CLEAR_APPLIES_TO.has(args.applyTo)) {
        throw invalid("clear_range.applyTo is invalid.");
      }
      return;
    case "update_sheet":
      assertOnlyKeys(args, ["action", "name", "newName"], operation);
      if (!SHEET_ACTIONS.has(args.action)) throw invalid("update_sheet.action is invalid.");
      assertString(args.name, "name", 1, 255);
      if (args.action === "rename") assertString(args.newName, "newName", 1, 255);
      return;
    case "update_workbook":
      assertOnlyKeys(args, ["calculationMode"], operation);
      if (args.calculationMode === undefined) {
        throw invalid("update_workbook requires calculationMode.");
      }
      if (!["Automatic", "Manual", "AutomaticExceptTables"].includes(args.calculationMode)) {
        throw invalid("update_workbook.calculationMode is invalid.");
      }
      return;
    case "copy_range_to":
      assertOnlyKeys(args, ["source", "destination", "type"], operation);
      assertPlainObject(args.source, "source");
      assertPlainObject(args.destination, "destination");
      assertRangeTarget(args.source, "source");
      assertRangeTarget(args.destination, "destination");
      if (args.type !== undefined && !["all", "values", "formulas", "formats"].includes(args.type)) {
        throw invalid("copy_range_to.type is invalid.");
      }
      return;
    case "read_range_image":
      assertOnlyKeys(args, ["sheet", "range"], operation);
      assertRangeTarget(args);
      return;
    case "read_sheets_metadata":
      assertNoUnexpectedArguments(args, operation);
      return;
    case "resize_range":
      assertOnlyKeys(args, ["sheet", "range", "columnWidth", "rowHeight", "autofitColumns", "autofitRows"], operation);
      assertRangeTarget(args);
      if (
        args.columnWidth === undefined && args.rowHeight === undefined &&
        args.autofitColumns !== true && args.autofitRows !== true
      ) {
        throw invalid("resize_range requires a width, height, or autofit option.");
      }
      assertOptionalPositiveNumber(args.columnWidth, "columnWidth");
      assertOptionalPositiveNumber(args.rowHeight, "rowHeight");
      assertOptionalBoolean(args.autofitColumns, "autofitColumns");
      assertOptionalBoolean(args.autofitRows, "autofitRows");
      return;
    case "update_sheet_view":
      assertOnlyKeys(args, ["sheet", "freezeRows", "freezeColumns", "showGridlines", "zoom", "selectRange"], operation);
      assertSheet(args);
      if (
        args.freezeRows === undefined && args.freezeColumns === undefined &&
        args.showGridlines === undefined && args.zoom === undefined && args.selectRange === undefined
      ) {
        throw invalid("update_sheet_view requires at least one view setting.");
      }
      assertOptionalNonnegativeInteger(args.freezeRows, "freezeRows");
      assertOptionalNonnegativeInteger(args.freezeColumns, "freezeColumns");
      assertOptionalBoolean(args.showGridlines, "showGridlines");
      if (args.zoom !== undefined && (!Number.isFinite(args.zoom) || args.zoom < 10 || args.zoom > 400)) {
        throw invalid("zoom must be a number from 10 through 400.");
      }
      if (args.selectRange !== undefined) assertA1Range(args.selectRange, "selectRange");
      return;
    case "format_range":
      assertOnlyKeys(args, ["sheet", "range", "format"], operation);
      assertRangeTarget(args);
      validateRangeFormat(args.format);
      return;
    case "chart":
      assertOnlyKeys(args, ["action", "sheet", "name", "type", "sourceRange", "title", "position"], operation);
      if (!CHART_ACTIONS.has(args.action)) throw invalid("chart.action is invalid.");
      assertSheet(args);
      if (args.action === "create") {
        assertString(args.name, "name", 1, 255);
        assertA1Range(args.sourceRange, "sourceRange");
        assertString(args.type, "type", 1, 96);
      } else {
        assertString(args.name, "name", 1, 255);
        if (args.action === "update" && args.title === undefined && args.sourceRange === undefined && args.position === undefined) {
          throw invalid("chart update requires title, sourceRange, or position.");
        }
      }
      if (args.title !== undefined) assertString(args.title, "title", 1, 512);
      if (args.sourceRange !== undefined) assertA1Range(args.sourceRange, "sourceRange");
      if (args.position !== undefined) validateChartPosition(args.position);
      return;
    case "table":
      assertOnlyKeys(args, ["action", "sheet", "name", "range", "hasHeaders", "rows", "index", "count"], operation);
      if (!TABLE_ACTIONS.has(args.action)) throw invalid("table.action is invalid.");
      assertSheet(args);
      assertString(args.name, "name", 1, 255);
      if (args.action === "create") {
        assertA1Range(args.range, "range");
        assertOptionalBoolean(args.hasHeaders, "hasHeaders");
      }
      if (args.action === "add_rows") assertMatrix(args.rows, "rows");
      if (args.action === "delete_rows") {
        assertOptionalNonnegativeInteger(args.index, "index");
        assertOptionalPositiveInteger(args.count, "count");
      }
      return;
    case "pivot_table":
      assertOnlyKeys(args, ["action", "sheet", "name", "source", "destination"], operation);
      if (!PIVOT_ACTIONS.has(args.action)) throw invalid("pivot_table.action is invalid.");
      assertSheet(args);
      assertString(args.name, "name", 1, 255);
      if (args.action === "create") {
        assertString(args.source, "source", 1, 512);
        assertA1Range(args.destination, "destination");
      }
      return;
    case "save":
      assertNoUnexpectedArguments(args, operation);
      return;
    default:
      throw excelLiveError("unsupported-operation", `Unsupported Excel Live operation: ${operation}.`);
  }
}

function assertSheet(args) {
  assertString(args.sheet, "sheet", 1, 255);
}

function assertRangeTarget(args, label = "args") {
  assertPlainObject(args, label);
  assertSheet(args);
  assertA1Range(args.range, "range");
}

function assertA1Range(value, label) {
  assertString(value, label, 1, 512);
  if (/\0|[\r\n]/u.test(value)) throw invalid(`${label} must be a single A1 range.`);
}

function assertIdentifier(value, label, maxLength) {
  assertString(value, label, 8, maxLength);
  if (!/^[A-Za-z0-9][A-Za-z0-9._:-]*$/u.test(value)) {
    throw invalid(`${label} contains unsupported characters.`);
  }
}

function assertMatrix(value, label, { stringsOnly = false } = {}) {
  assertArray(value, label, 1, MAX_MATRIX_CELLS);
  let width = null;
  let cells = 0;
  for (const row of value) {
    assertArray(row, `${label} row`, 1, MAX_MATRIX_CELLS);
    if (width == null) width = row.length;
    if (row.length !== width) throw invalid(`${label} must be rectangular.`);
    cells += row.length;
    if (cells > MAX_MATRIX_CELLS) throw invalid(`${label} exceeds ${MAX_MATRIX_CELLS} cells.`);
    for (const cell of row) {
      if (stringsOnly && typeof cell !== "string") {
        throw invalid(`${label} entries must be strings.`);
      }
      assertJsonValue(cell, `${label} cell`);
    }
  }
}

function assertMatchingMatrices(...matrices) {
  const present = matrices.filter((matrix) => matrix !== undefined);
  if (present.length < 2) return;
  const [first] = present;
  for (const matrix of present.slice(1)) {
    if (matrix.length !== first.length || matrix.some((row, index) => row.length !== first[index].length)) {
      throw invalid("values, formulas, and numberFormat must have matching rectangular dimensions.");
    }
  }
}

function assertNoUnexpectedArguments(args, operation) {
  if (Object.keys(args).length > 0) throw invalid(`${operation} does not accept arguments.`);
}

function assertOnlyKeys(args, allowed, operation) {
  for (const key of Object.keys(args)) {
    if (!allowed.includes(key)) throw invalid(`${operation}.${key} is not supported.`);
  }
}

function validateSearchOptions(options) {
  assertPlainObject(options, "options");
  assertOnlyKeys(options, ["matchCase", "completeMatch", "maxResults"], "search_workbook.options");
  assertOptionalBoolean(options.matchCase, "options.matchCase");
  assertOptionalBoolean(options.completeMatch, "options.completeMatch");
  if (options.maxResults !== undefined && (!Number.isSafeInteger(options.maxResults) || options.maxResults < 1 || options.maxResults > MAX_SEARCH_RESULTS)) {
    throw invalid(`options.maxResults must be an integer from 1 through ${MAX_SEARCH_RESULTS}.`);
  }
}

function validateRangeFormat(format) {
  assertPlainObject(format, "format");
  assertOnlyKeys(format, [
    "numberFormat",
    "fill",
    "font",
    "horizontalAlignment",
    "verticalAlignment",
    "wrapText",
    "columnWidth",
    "rowHeight",
    "indentLevel",
    "textOrientation",
    "borders",
  ], "format");
  if (Object.keys(format).length === 0) throw invalid("format cannot be empty.");
  if (format.numberFormat !== undefined) assertMatrix(format.numberFormat, "format.numberFormat", { stringsOnly: true });
  if (format.fill !== undefined) {
    assertPlainObject(format.fill, "format.fill");
    assertOnlyKeys(format.fill, ["color"], "format.fill");
    assertString(format.fill.color, "format.fill.color", 1, 64);
  }
  if (format.font !== undefined) {
    assertPlainObject(format.font, "format.font");
    assertOnlyKeys(format.font, ["bold", "italic", "underline", "strikethrough", "color", "name", "size"], "format.font");
    for (const key of ["bold", "italic", "strikethrough"]) assertOptionalBoolean(format.font[key], `format.font.${key}`);
    if (format.font.underline !== undefined) assertString(format.font.underline, "format.font.underline", 1, 64);
    if (format.font.color !== undefined) assertString(format.font.color, "format.font.color", 1, 64);
    if (format.font.name !== undefined) assertString(format.font.name, "format.font.name", 1, 255);
    assertOptionalPositiveNumber(format.font.size, "format.font.size");
  }
  for (const key of ["horizontalAlignment", "verticalAlignment"]) {
    if (format[key] !== undefined) assertString(format[key], `format.${key}`, 1, 64);
  }
  assertOptionalBoolean(format.wrapText, "format.wrapText");
  assertOptionalPositiveNumber(format.columnWidth, "format.columnWidth");
  assertOptionalPositiveNumber(format.rowHeight, "format.rowHeight");
  assertOptionalNonnegativeInteger(format.indentLevel, "format.indentLevel");
  if (format.textOrientation !== undefined && (!Number.isSafeInteger(format.textOrientation) || format.textOrientation < -90 || format.textOrientation > 90)) {
    throw invalid("format.textOrientation must be an integer from -90 through 90.");
  }
  if (format.borders !== undefined) {
    assertArray(format.borders, "format.borders", 1, 16);
    for (const border of format.borders) {
      assertPlainObject(border, "format.borders entry");
      assertOnlyKeys(border, ["side", "style", "color"], "format.borders entry");
      assertString(border.side, "format.borders side", 1, 64);
      if (border.style !== undefined) assertString(border.style, "format.borders style", 1, 64);
      if (border.color !== undefined) assertString(border.color, "format.borders color", 1, 64);
    }
  }
}

function validateChartPosition(position) {
  assertPlainObject(position, "position");
  assertOnlyKeys(position, ["start", "end"], "position");
  assertA1Range(position.start, "position.start");
  assertA1Range(position.end, "position.end");
}

function assertPlainObject(value, label) {
  if (value == null || typeof value !== "object" || Array.isArray(value)) {
    throw invalid(`${label} must be an object.`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw invalid(`${label} must be a plain object.`);
  }
}

function assertArray(value, label, minimum, maximum) {
  if (!Array.isArray(value) || value.length < minimum || value.length > maximum) {
    throw invalid(`${label} must contain ${minimum} through ${maximum} entries.`);
  }
}

function assertString(value, label, minimum, maximum) {
  if (typeof value !== "string" || value.length < minimum || value.length > maximum) {
    throw invalid(`${label} must be a string from ${minimum} through ${maximum} characters.`);
  }
}

function assertExactNumber(value, expected, label) {
  if (value !== expected) throw invalid(`${label} must equal ${expected}.`);
}

function assertOptionalPositiveNumber(value, label) {
  if (value !== undefined && (!Number.isFinite(value) || value <= 0)) {
    throw invalid(`${label} must be a positive number.`);
  }
}

function assertOptionalPositiveInteger(value, label) {
  if (value !== undefined && (!Number.isSafeInteger(value) || value <= 0)) {
    throw invalid(`${label} must be a positive integer.`);
  }
}

function assertOptionalNonnegativeInteger(value, label) {
  if (value !== undefined && (!Number.isSafeInteger(value) || value < 0)) {
    throw invalid(`${label} must be a non-negative integer.`);
  }
}

function assertOptionalBoolean(value, label) {
  if (value !== undefined && typeof value !== "boolean") throw invalid(`${label} must be boolean.`);
}

function assertJsonValue(value, label, depth = 0) {
  if (depth > 20) throw invalid(`${label} is nested too deeply.`);
  if (value == null || typeof value === "string" || typeof value === "boolean") return;
  if (typeof value === "number" && Number.isFinite(value)) return;
  if (Array.isArray(value)) {
    for (const item of value) assertJsonValue(item, label, depth + 1);
    return;
  }
  if (typeof value === "object") {
    assertPlainObject(value, label);
    for (const [key, item] of Object.entries(value)) {
      assertString(key, `${label} key`, 1, 256);
      assertJsonValue(item, label, depth + 1);
    }
    return;
  }
  throw invalid(`${label} must be JSON serializable.`);
}

function cloneJson(value, label) {
  assertJsonValue(value, label);
  const serialized = JSON.stringify(value);
  if (Buffer.byteLength(serialized) > MAX_REQUEST_BYTES) {
    throw excelLiveError("request-too-large", `Excel request exceeds ${MAX_REQUEST_BYTES} bytes.`);
  }
  return JSON.parse(serialized);
}

function invalid(message) {
  return excelLiveError("invalid-request", message);
}
