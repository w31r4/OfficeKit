import { create } from "@bufbuild/protobuf";

import { Workbook } from "../spreadsheet/index.mjs";
import {
  ArtifactFamily,
  CellFormulaKind,
  CodecOperation,
  SpreadsheetCalculationMode,
  SpreadsheetWorksheetVisibility,
  WorkbookDateSystem,
} from "../generated/office_kit/artifact/v1/office_artifact_pb.js";
import { FileBlob } from "../shared/file-blob.mjs";
import { normalizeDataBarConfig, normalizeIconSetConfig } from "../spreadsheet/conditional-formats.mjs";
import { normalizeSpreadsheetDataValidationRule } from "../spreadsheet/data-validations.mjs";
import { XLSX_THEME_COLOR_NAMES, normalizeXlsxStyle, normalizeXlsxThemeConfig } from "../spreadsheet/ooxml-styles.mjs";
import { deterministicSpreadsheetGuid } from "../spreadsheet/ooxml-threaded-comments.mjs";
import { OfficeKitCodecError } from "./office-kit-error.mjs";
import {
  assertCodecOptions,
  codecLimits,
  inputBytes,
  invokeOfficeKit,
  OFFICE_KIT_PROTOCOL_VERSION,
} from "./office-kit-runtime.mjs";
import { assertTrustedImportedState } from "./office-kit-source-state.mjs";
import { spreadsheetChartFromWire, spreadsheetChartSnapshot, wireWorksheetCharts } from "./office-kit-spreadsheet-charts.mjs";
import { hydrateWorksheetDataTable, wireWorksheetDataTables } from "./office-kit-spreadsheet-data-tables.mjs";
import { spreadsheetImageFromWire, spreadsheetImageSnapshot, wireWorksheetImages } from "./office-kit-spreadsheet-images.mjs";
import { hydrateWorkbookPivots, wireWorksheetPivots } from "./office-kit-spreadsheet-pivots.mjs";
import { publicWorksheetProtectionFromWire, wireWorksheetProtection, worksheetProtectionPublicSnapshot } from "./office-kit-spreadsheet-protection.mjs";
import { spreadsheetSparklineFromWire, spreadsheetSparklineSnapshot, wireWorksheetSparklines } from "./office-kit-spreadsheet-sparklines.mjs";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const WORKBOOK_STATE = Symbol.for("office-kit.workbook-state");
const TABLE_STATE = Symbol.for("office-kit.table-state");

const MAX_XLSX_NUMBER_FORMAT_CODE_LENGTH = 4096;
const MAX_XLSX_FORMULA_LENGTH = 8192;
const MAX_XLSX_FORMULA_TOPOLOGY_CELLS = 1_048_576;
const XLSX_FORMULA_METADATA_KEYS = new Set([
  "formulaType", "sharedIndex", "sharedRef", "arrayRef", "dynamicArrayRef",
  "spillParent", "spillAnchor", "spillRange", "spillValues", "spillError",
]);
const XLSX_NUMBER_FORMAT_STYLE_KEYS = new Set(["numberFormat", "numFmt"]);
const EXCEL_ERRORS = new Set(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#UNKNOWN!", "#CONNECT!", "#CYCLE!"]);
const XLSX_THEME_WIRE_FIELDS = [
  ["dk1", "dk1Rgb"], ["lt1", "lt1Rgb"], ["dk2", "dk2Rgb"], ["lt2", "lt2Rgb"],
  ["accent1", "accent1Rgb"], ["accent2", "accent2Rgb"], ["accent3", "accent3Rgb"],
  ["accent4", "accent4Rgb"], ["accent5", "accent5Rgb"], ["accent6", "accent6Rgb"],
  ["hlink", "hlinkRgb"], ["folHlink", "folHlinkRgb"],
];

function publicWorksheetVisibility(value) {
  if (value === SpreadsheetWorksheetVisibility.HIDDEN) return "hidden";
  if (value === SpreadsheetWorksheetVisibility.VERY_HIDDEN) return "veryHidden";
  return "visible";
}

function wireWorksheetVisibility(value) {
  if (value === "hidden") return SpreadsheetWorksheetVisibility.HIDDEN;
  if (value === "veryHidden") return SpreadsheetWorksheetVisibility.VERY_HIDDEN;
  if (value === "visible") return SpreadsheetWorksheetVisibility.VISIBLE;
  throw new OfficeKitCodecError(`Unsupported worksheet visibility ${value}; expected visible, hidden, or veryHidden.`, [], { code: "invalid_worksheet_visibility" });
}

function worksheetMetadataSnapshot(sheet) {
  return { name: sheet.name, visibility: sheet.visibility };
}

function wireWorksheetMetadata(sheet, slot) {
  const unchanged = slot && JSON.stringify(worksheetMetadataSnapshot(sheet)) === JSON.stringify(slot.publicSnapshot);
  return {
    visibility: unchanged ? slot.wire.visibility : wireWorksheetVisibility(sheet.visibility),
    source: slot?.wire.source,
  };
}

function wireWorksheetFreezePane(sheet, slot) {
  const rows = sheet.freezePanes?.rows || 0;
  const columns = sheet.freezePanes?.columns || 0;
  // An omitted native pane is observably different from a zero-valued pane.
  // Keep omission for an untouched imported sheet, while a source-bound pane
  // can still be cleared by explicitly sending the zero-valued replacement.
  if (rows === 0 && columns === 0 && slot?.wire?.freezePane == null) return undefined;
  return {
    rows,
    columns,
    topLeftCell: sheet.freezePanes?.topLeftCell || "",
    activePane: sheet.freezePanes?.activePane || "",
  };
}

function workbookViewSnapshots(workbook) {
  return workbook.windows.items.map((window) => ({
    activeWorksheetId: window.getActiveWorksheet().id,
    selectedWorksheetIds: window.getSelectedWorksheets().map((sheet) => sheet.id),
  }));
}

function wireWorkbookViews(workbook, state) {
  const slots = state?.viewSlots || [];
  if (state && slots.length === 0 && !workbook._activeWorksheetId && workbook.windows.count === 1)
    return { view: undefined, additionalViews: [] };
  const snapshots = workbookViewSnapshots(workbook);
  const views = snapshots.map((snapshot, index) => {
    const slot = slots[index];
    if (slot && JSON.stringify(snapshot) === JSON.stringify(slot.publicSnapshot)) return slot.wire;
    return { ...snapshot, source: slot?.wire.source };
  });
  return { view: views[0], additionalViews: views.slice(1) };
}

function cellCoordinates(address) {
  const match = /^([A-Z]{1,3})([1-9]\d*)$/i.exec(String(address));
  if (!match) throw new OfficeKitCodecError(`Cell address ${address} is not valid A1 notation.`, [], { code: "invalid_cell_address" });
  let column = 0;
  for (const character of match[1].toUpperCase()) column = column * 26 + character.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, column: column - 1 };
}

function cellAddress(row, column) {
  let number = Number(column) + 1;
  let label = "";
  while (number > 0) {
    number -= 1;
    label = String.fromCharCode(65 + number % 26) + label;
    number = Math.floor(number / 26);
  }
  return `${label}${Number(row) + 1}`;
}

function formulaRangeBounds(reference, location) {
  const pieces = String(reference || "").split(":");
  if (pieces.length < 1 || pieces.length > 2 || !pieces[0]) throw new OfficeKitCodecError(`Cell ${location} formula reference ${reference || "(empty)"} is not a bounded A1 range.`, [], { code: "invalid_cell_formula" });
  let first;
  let second;
  try {
    first = cellCoordinates(pieces[0].replaceAll("$", ""));
    second = pieces[1] ? cellCoordinates(pieces[1].replaceAll("$", "")) : first;
  } catch {
    throw new OfficeKitCodecError(`Cell ${location} formula reference ${reference} is invalid.`, [], { code: "invalid_cell_formula" });
  }
  for (const coordinate of [first, second]) {
    if (coordinate.row >= 1_048_576 || coordinate.column >= 16_384) throw new OfficeKitCodecError(`Cell ${location} formula reference ${reference} exceeds XLSX limits.`, [], { code: "invalid_cell_formula" });
  }
  if (first.row > second.row || first.column > second.column) throw new OfficeKitCodecError(`Cell ${location} formula reference ${reference} must be top-left to bottom-right.`, [], { code: "invalid_cell_formula" });
  return { top: first.row, left: first.column, bottom: second.row, right: second.column, cellCount: (second.row - first.row + 1) * (second.column - first.column + 1) };
}

function partialSharedFormulaRanges(diagnostics) {
  const byWorksheetName = new Map();
  for (const diagnostic of diagnostics || []) {
    if (diagnostic?.code !== "partial_shared_formula_preserved") continue;
    const worksheetName = String(diagnostic.sourcePath || "");
    const reference = String(diagnostic.sourceIdentity || "");
    if (!worksheetName || !reference) {
      throw new OfficeKitCodecError("OfficeKit returned a partial shared-formula diagnostic without a worksheet name and range identity.", [], { code: "invalid_office_kit_diagnostic" });
    }
    const bounds = formulaRangeBounds(reference, `${worksheetName} partial shared formula`);
    const ranges = byWorksheetName.get(worksheetName) || [];
    ranges.push({ reference, bounds });
    byWorksheetName.set(worksheetName, ranges);
  }
  return byWorksheetName;
}

function normalizedFormula(value) {
  const formula = String(value || "");
  return formula && !formula.startsWith("=") ? `=${formula}` : formula;
}

function validateFormulaText(value, location, required = false) {
  const formula = normalizedFormula(value);
  const body = formula.startsWith("=") ? formula.slice(1) : formula;
  if (required && !body.trim()) throw new OfficeKitCodecError(`Cell ${location} requires non-empty formula text.`, [], { code: "invalid_cell_formula" });
  if (body.length > MAX_XLSX_FORMULA_LENGTH || /\p{Cc}/u.test(body)) throw new OfficeKitCodecError(`Cell ${location} formula is outside the bounded XLSX formula profile.`, [], { code: "invalid_cell_formula" });
  return formula;
}

function translateSharedFormula(value, source, target) {
  const formula = normalizedFormula(value);
  const rowOffset = target.row - source.row;
  const columnOffset = target.column - source.column;
  const protectedParts = [];
  const protectedFormula = formula.replace(/"(?:[^"]|"")*"|\[[^\]]*\]/g, (part) => {
    const token = `\uE000${protectedParts.length}\uE001`;
    protectedParts.push(part);
    return token;
  });
  const shifted = protectedFormula.replace(/(?<![A-Za-z0-9_.])(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_. ]*))!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_])/g, (match, quotedSheet, bareSheet, absoluteColumn, columnText, absoluteRow, rowText, offset, sourceText) => {
    if (/^\s*\(/.test(sourceText.slice(offset + match.length))) return match;
    const coordinate = cellCoordinates(`${columnText}${rowText}`);
    const column = absoluteColumn ? coordinate.column : coordinate.column + columnOffset;
    const row = absoluteRow ? coordinate.row : coordinate.row + rowOffset;
    const prefix = quotedSheet != null ? `'${quotedSheet}'!` : bareSheet != null ? `${bareSheet}!` : "";
    if (column < 0 || column >= 16_384 || row < 0 || row >= 1_048_576) return `${prefix}#REF!`;
    return `${prefix}${absoluteColumn || ""}${cellAddress(row, column).replace(/\d+$/, `${absoluteRow || ""}${row + 1}`)}`;
  });
  return shifted.replace(/\uE000(\d+)\uE001/g, (_match, index) => protectedParts[Number(index)] || "");
}

function cellFormulaMetadata(address, cell) {
  const location = address;
  const type = cell.formulaType == null ? "" : String(cell.formulaType);
  if (!type && cell.formula && cell.spillError) throw new OfficeKitCodecError(`Cell ${location} blocked dynamic array cannot be exported through the current bounded OfficeKit slice.`, [], { code: "unsupported_dynamic_array_edit" });
  const inferredDynamic = !type && Boolean(cell.formula) && Boolean(cell.spillRange) && !cell.spillError;
  if (!type && !inferredDynamic && [cell.sharedIndex, cell.sharedRef, cell.arrayRef, cell.dynamicArrayRef].every((value) => value == null || value === "")) return undefined;
  if (type === "shared") {
    if (!Number.isInteger(cell.sharedIndex) || cell.sharedIndex < 0 || cell.sharedIndex > 0xffff_ffff) throw new OfficeKitCodecError(`Cell ${location} shared formula requires an unsigned sharedIndex.`, [], { code: "invalid_cell_formula" });
    const reference = String(cell.sharedRef || "");
    formulaRangeBounds(reference, location);
    validateFormulaText(cell.formula, location, true);
    if (cell.arrayRef != null) throw new OfficeKitCodecError(`Cell ${location} shared formula must not set arrayRef.`, [], { code: "invalid_cell_formula" });
    return { kind: CellFormulaKind.SHARED, sharedIndex: cell.sharedIndex, reference };
  }
  if (type === "array") {
    const reference = String(cell.arrayRef || "");
    formulaRangeBounds(reference, location);
    validateFormulaText(cell.formula, location, true);
    if (cell.sharedIndex != null || cell.sharedRef != null) throw new OfficeKitCodecError(`Cell ${location} legacy array formula must not set shared metadata.`, [], { code: "invalid_cell_formula" });
    return { kind: CellFormulaKind.ARRAY, sharedIndex: 0, reference };
  }
  if (type === "dynamicArray" || inferredDynamic) {
    const reference = String(cell.dynamicArrayRef || cell.spillRange || "");
    const bounds = formulaRangeBounds(reference, location);
    const anchor = cellCoordinates(address);
    if (anchor.row !== bounds.top || anchor.column !== bounds.left) throw new OfficeKitCodecError(`Cell ${location} dynamic array formula must be the top-left anchor of ${reference}.`, [], { code: "invalid_cell_formula" });
    validateFormulaText(cell.formula, location, true);
    if (cell.spillError) throw new OfficeKitCodecError(`Cell ${location} blocked dynamic array cannot be exported through the current bounded OfficeKit slice.`, [], { code: "unsupported_dynamic_array_edit" });
    if (cell.sharedIndex != null || cell.sharedRef != null || cell.arrayRef != null) throw new OfficeKitCodecError(`Cell ${location} dynamic array formula must not set shared or legacy-array metadata.`, [], { code: "invalid_cell_formula" });
    return { kind: CellFormulaKind.DYNAMIC_ARRAY, sharedIndex: 0, reference };
  }
  throw new OfficeKitCodecError(`Cell ${location} formula type ${type || "unspecified"} is outside the OfficeKit XLSX formula slice.`, [], { code: "unsupported_cell_formula" });
}

function validateFormulaTopology(cells, sheetName) {
  const byCoordinate = new Map(cells.map((cell) => [`${cell.row}:${cell.column}`, cell]));
  if (byCoordinate.size !== cells.length) throw new OfficeKitCodecError(`Worksheet ${sheetName} contains duplicate cell coordinates.`, [], { code: "duplicate_cell" });
  const sharedGroups = new Map();
  for (const cell of cells) {
    validateFormulaText(cell.formula, `${sheetName}!${cellAddress(cell.row, cell.column)}`, Boolean(cell.formulaMetadata) && cell.formulaMetadata.kind !== CellFormulaKind.DATA_TABLE);
    if (cell.formulaMetadata?.kind === CellFormulaKind.SHARED) {
      const key = cell.formulaMetadata.sharedIndex;
      if (!sharedGroups.has(key)) sharedGroups.set(key, []);
      sharedGroups.get(key).push(cell);
    }
  }
  for (const [index, members] of sharedGroups) {
    const references = new Set(members.map((cell) => cell.formulaMetadata.reference.toUpperCase()));
    if (references.size !== 1) throw new OfficeKitCodecError(`Worksheet ${sheetName} shared formula si=${index} has inconsistent references.`, [], { code: "invalid_cell_formula" });
    const reference = members[0].formulaMetadata.reference;
    const bounds = formulaRangeBounds(reference, `${sheetName}!${cellAddress(members[0].row, members[0].column)}`);
    const memberMap = new Map(members.map((cell) => [`${cell.row}:${cell.column}`, cell]));
    const expectedCount = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1);
    if (memberMap.size !== expectedCount) throw new OfficeKitCodecError(`Worksheet ${sheetName} shared formula si=${index} declares ${reference} with ${expectedCount} cells but contains ${memberMap.size} members.`, [], { code: "invalid_cell_formula" });
    const master = memberMap.get(`${bounds.top}:${bounds.left}`);
    if (!master) throw new OfficeKitCodecError(`Worksheet ${sheetName} shared formula si=${index} is missing its top-left master.`, [], { code: "invalid_cell_formula" });
    for (let row = bounds.top; row <= bounds.bottom; row += 1) {
      for (let column = bounds.left; column <= bounds.right; column += 1) {
        const member = memberMap.get(`${row}:${column}`);
        if (!member) throw new OfficeKitCodecError(`Worksheet ${sheetName} shared formula si=${index} is missing ${cellAddress(row, column)}.`, [], { code: "invalid_cell_formula" });
        const expected = translateSharedFormula(master.formula, { row: bounds.top, column: bounds.left }, { row, column });
        if (normalizedFormula(member.formula) !== expected) throw new OfficeKitCodecError(`Cell ${sheetName}!${cellAddress(row, column)} expanded shared formula must be ${expected}.`, [], { code: "invalid_cell_formula" });
      }
    }
  }
  const occupied = new Map();
  const sharedRoots = [...sharedGroups.values()].map((members) => members[0]);
  const topologyRoots = [...sharedRoots, ...cells.filter((item) => [CellFormulaKind.ARRAY, CellFormulaKind.DYNAMIC_ARRAY, CellFormulaKind.DATA_TABLE].includes(item.formulaMetadata?.kind))];
  let topologyCellCount = 0;
  for (const cell of topologyRoots) {
    const metadata = cell.formulaMetadata;
    const bounds = formulaRangeBounds(metadata.reference, `${sheetName}!${cellAddress(cell.row, cell.column)}`);
    topologyCellCount += bounds.cellCount;
    if (topologyCellCount > MAX_XLSX_FORMULA_TOPOLOGY_CELLS) throw new OfficeKitCodecError(`Cell ${sheetName}!${cellAddress(cell.row, cell.column)} native formula topology exceeds ${MAX_XLSX_FORMULA_TOPOLOGY_CELLS} cells.`, [], { code: "invalid_cell_formula" });
    const dynamic = metadata.kind === CellFormulaKind.DYNAMIC_ARRAY;
    const array = metadata.kind === CellFormulaKind.ARRAY || dynamic;
    const dataTable = metadata.kind === CellFormulaKind.DATA_TABLE;
    const owner = metadata.kind === CellFormulaKind.SHARED ? `shared:${metadata.sharedIndex}` : `${dataTable ? "data-table" : dynamic ? "dynamic" : "array"}:${cell.row}:${cell.column}`;
    if ((array || dataTable) && (cell.row !== bounds.top || cell.column !== bounds.left)) throw new OfficeKitCodecError(`Cell ${sheetName}!${cellAddress(cell.row, cell.column)} ${dataTable ? "data table" : `${dynamic ? "dynamic" : "legacy"} array`} formula must be the top-left anchor of ${metadata.reference}.`, [], { code: "invalid_cell_formula" });
    for (let row = bounds.top; row <= bounds.bottom; row += 1) {
      for (let column = bounds.left; column <= bounds.right; column += 1) {
        const key = `${row}:${column}`;
        if (occupied.has(key) && occupied.get(key) !== owner) throw new OfficeKitCodecError(`Cell ${sheetName}!${cellAddress(cell.row, cell.column)} formula range ${metadata.reference} overlaps another native formula range.`, [], { code: "invalid_cell_formula" });
        occupied.set(key, owner);
        const nested = byCoordinate.get(key);
        if ((array || dataTable) && (row !== cell.row || column !== cell.column) && nested?.formula) throw new OfficeKitCodecError(`Cell ${sheetName}!${cellAddress(row, column)} must not contain another formula inside ${dataTable ? "data table" : `${dynamic ? "dynamic" : "legacy"} array`} range ${metadata.reference}.`, [], { code: "invalid_cell_formula" });
      }
    }
  }
}

function itemCount(collection) {
  return Array.isArray(collection?.items) ? collection.items.length : 0;
}

function numberFormatCode(value, address) {
  if (value == null || value === "") return "";
  if (typeof value !== "string") throw new OfficeKitCodecError(`Cell ${address} number format must be a string.`, [], { code: "invalid_cell_number_format" });
  if (/^general$/i.test(value)) return "";
  if (value.length > MAX_XLSX_NUMBER_FORMAT_CODE_LENGTH) throw new OfficeKitCodecError(`Cell ${address} number format exceeds ${MAX_XLSX_NUMBER_FORMAT_CODE_LENGTH} characters.`, [], { code: "invalid_cell_number_format" });
  if (/\p{Cc}/u.test(value)) throw new OfficeKitCodecError(`Cell ${address} number format contains a control character.`, [], { code: "invalid_cell_number_format" });
  return value;
}

function cellNumberFormatCode(cell, address) {
  return numberFormatCode(cell?.style?.numberFormat ?? cell?.style?.numFmt, address);
}

function invalidCellStyle(address, message, cause) {
  throw new OfficeKitCodecError(`Cell ${address} ${message}`, [], { code: "invalid_cell_style", cause });
}

function wireSpreadsheetColor(value, address, component) {
  if (value == null) return undefined;
  if (typeof value === "string") {
    const rgb = value.replace(/^#/, "").slice(-6).toUpperCase();
    if (!/^[0-9A-F]{6}$/.test(rgb)) invalidCellStyle(address, `${component} color must be six/eight-digit RGB or a supported symbolic color.`);
    return { source: { case: "rgb", value: rgb } };
  }
  const tint = value.tint == null ? undefined : Number(value.tint);
  if (value.theme != null) return { source: { case: "theme", value: Number(value.theme) }, tint };
  if (value.indexed != null) return { source: { case: "indexed", value: Number(value.indexed) }, tint };
  if (value.auto === true) return { source: { case: "automatic", value: true }, tint };
  if (value.rgb != null) return { ...wireSpreadsheetColor(value.rgb, address, component), tint };
  invalidCellStyle(address, `${component} color has no supported source.`);
}

function wireBorderEdge(edge, address, name) {
  if (!edge?.style) return undefined;
  return { style: String(edge.style), color: wireSpreadsheetColor(edge.color, address, `${name} border`) };
}

function wireCellStyle(style, address) {
  const keys = Object.keys(style || {}).filter((key) => style[key] != null && !XLSX_NUMBER_FORMAT_STYLE_KEYS.has(key));
  if (keys.length === 0) return undefined;
  let normalized;
  try {
    normalized = normalizeXlsxStyle(style);
  } catch (cause) {
    invalidCellStyle(address, `has invalid static formatting: ${cause.message}`, cause);
  }
  const font = normalized.font;
  const fontInput = style?.font || {};
  const fontField = (flat, nested = flat) => style?.[flat] != null || fontInput[nested] != null;
  const hasFontColor = style?.fontColor != null || style?.color != null || fontInput.color != null;
  const hasFont = [
    fontField("bold"), fontField("italic"), fontField("underline"), fontField("strike"),
    fontField("fontSize", "size"), fontField("fontFamily", "name"), hasFontColor,
  ].some(Boolean);
  const fill = typeof normalized.fill === "string" ? { patternType: "solid", foreground: normalized.fill } : normalized.fill;
  const border = normalized.border;
  const uniformEdge = border?.style ? { style: border.style, color: border.color } : undefined;
  // The public shorthand `{ border: { style, color } }` means the outside
  // perimeter. It must not silently create diagonal or interior borders when
  // a source-bound workbook is exported again.
  const edge = (name) => wireBorderEdge(
    (name === "left" || name === "right" || name === "top" || name === "bottom" ? uniformEdge : undefined) || border?.[name],
    address,
    name,
  );
  return {
    // `CellStyleArtifact` is sparse. Preserve that field presence instead of
    // converting a supplied color/bold flag into unrelated Aptos/11/false
    // defaults, which would make an untouched imported differential style a
    // semantic edit.
    font: hasFont ? {
      bold: fontField("bold") ? font.bold : undefined,
      italic: fontField("italic") ? font.italic : undefined,
      underline: fontField("underline") ? String(font.underline === true ? "single" : font.underline || "none") : undefined,
      strike: fontField("strike") ? font.strike : undefined,
      color: hasFontColor ? wireSpreadsheetColor(font.color, address, "font") : undefined,
      sizePoints: fontField("fontSize", "size") ? font.size : undefined,
      name: fontField("fontFamily", "name") ? font.name : undefined,
    } : undefined,
    fill: fill ? {
      patternType: fill.patternType,
      foreground: wireSpreadsheetColor(fill.foreground, address, "fill foreground"),
      background: wireSpreadsheetColor(fill.background, address, "fill background"),
    } : undefined,
    border: border ? {
      left: edge("left"), right: edge("right"), top: edge("top"), bottom: edge("bottom"),
      diagonal: edge("diagonal"), start: edge("start"), end: edge("end"),
      horizontal: edge("horizontal"), vertical: edge("vertical"),
      diagonalUp: border.diagonalUp,
      diagonalDown: border.diagonalDown,
      outline: border.outline,
    } : undefined,
    alignment: normalized.alignment ? {
      horizontal: normalized.alignment.horizontal,
      vertical: normalized.alignment.vertical,
      wrapText: normalized.alignment.wrapText,
      textRotation: normalized.alignment.textRotation,
      indent: normalized.alignment.indent,
      shrinkToFit: normalized.alignment.shrinkToFit,
      readingOrder: normalized.alignment.readingOrder,
    } : undefined,
    protection: normalized.protection ? {
      locked: normalized.protection.locked,
      hidden: normalized.protection.hidden,
    } : undefined,
  };
}

function spreadsheetColorFromWire(color) {
  if (!color?.source?.case) return undefined;
  const tint = color.tint == null || color.tint === 0 ? {} : { tint: color.tint };
  if (color.source.case === "rgb") return color.tint == null || color.tint === 0
    ? `#${String(color.source.value).slice(-6).toUpperCase()}`
    : { rgb: `#${String(color.source.value).slice(-6).toUpperCase()}`, ...tint };
  if (color.source.case === "theme") return { theme: color.source.value, ...tint };
  if (color.source.case === "indexed") return { indexed: color.source.value, ...tint };
  if (color.source.case === "automatic") return { auto: true, ...tint };
  return undefined;
}

function borderEdgeFromWire(edge) {
  if (!edge?.style) return undefined;
  return { style: edge.style, color: spreadsheetColorFromWire(edge.color) || "#000000" };
}

function cellStyleFromWire(source) {
  if (!source) return undefined;
  const style = {};
  if (source.font) {
    style.font = {
      ...(source.font.bold == null ? {} : { bold: source.font.bold }),
      ...(source.font.italic == null ? {} : { italic: source.font.italic }),
      ...(source.font.underline == null ? {} : { underline: source.font.underline }),
      ...(source.font.strike == null ? {} : { strike: source.font.strike }),
      ...(source.font.color ? { color: spreadsheetColorFromWire(source.font.color) } : {}),
      ...(source.font.sizePoints == null ? {} : { size: source.font.sizePoints }),
      ...(source.font.name == null ? {} : { name: source.font.name }),
    };
  }
  if (source.fill) {
    const foreground = spreadsheetColorFromWire(source.fill.foreground);
    const background = spreadsheetColorFromWire(source.fill.background);
    style.fill = source.fill.patternType === "solid" && typeof foreground === "string" && !background
      ? foreground
      : { patternType: source.fill.patternType || "none", ...(foreground ? { foreground } : {}), ...(background ? { background } : {}) };
  }
  if (source.border) {
    const border = {};
    for (const name of ["left", "right", "top", "bottom", "diagonal", "start", "end", "horizontal", "vertical"]) {
      const value = borderEdgeFromWire(source.border[name]);
      if (value) border[name] = value;
    }
    for (const [wire, model] of [["diagonalUp", "diagonalUp"], ["diagonalDown", "diagonalDown"], ["outline", "outline"]]) {
      if (source.border[wire] != null) border[model] = source.border[wire];
    }
    const perimeter = [border.left, border.right, border.top, border.bottom];
    const samePerimeter = perimeter.every(Boolean) && perimeter.every((candidate) => JSON.stringify(candidate) === JSON.stringify(perimeter[0]));
    const hasExtras = border.diagonal || border.start || border.end || border.horizontal || border.vertical || border.diagonalUp != null || border.diagonalDown != null || border.outline != null;
    style.border = samePerimeter && !hasExtras ? perimeter[0] : border;
  }
  if (source.alignment) {
    style.alignment = Object.fromEntries(Object.entries({
      horizontal: source.alignment.horizontal,
      vertical: source.alignment.vertical,
      wrapText: source.alignment.wrapText,
      textRotation: source.alignment.textRotation,
      indent: source.alignment.indent,
      shrinkToFit: source.alignment.shrinkToFit,
      readingOrder: source.alignment.readingOrder,
    }).filter(([, value]) => value != null));
  }
  if (source.protection) {
    style.protection = Object.fromEntries(Object.entries({ locked: source.protection.locked, hidden: source.protection.hidden }).filter(([, value]) => value != null));
  }
  return Object.keys(style).length ? style : undefined;
}

function dynamicArrayCellSnapshot(cell) {
  return {
    formula: cell.formula == null ? null : String(cell.formula),
    formulaType: cell.formulaType == null ? null : String(cell.formulaType),
    dynamicArrayRef: cell.dynamicArrayRef == null ? null : String(cell.dynamicArrayRef),
    spillRange: cell.spillRange == null ? null : String(cell.spillRange),
    spillParent: cell.spillParent == null ? null : String(cell.spillParent),
    spillAnchor: cell.spillAnchor == null ? null : String(cell.spillAnchor),
    spillError: cell.spillError == null ? null : String(cell.spillError),
  };
}

function dynamicArrayRangeSnapshot(sheet, address, reference) {
  const bounds = formulaRangeBounds(reference, `${sheet.name}!${address}`);
  const output = [];
  for (let row = bounds.top; row <= bounds.bottom; row += 1) {
    for (let column = bounds.left; column <= bounds.right; column += 1) {
      const candidateAddress = cellAddress(row, column);
      const cell = sheet.store.get(candidateAddress);
      output.push({
        address: candidateAddress,
        value: cell.value instanceof Date ? { type: "date", value: cell.value.getTime() } : cell.value ?? null,
        ...dynamicArrayCellSnapshot(cell),
      });
    }
  }
  return output;
}

function dynamicArrayRecalculationMatches(sheet, address, slot) {
  const anchor = slot.cell;
  const reference = String(slot.publicSnapshot.dynamicArrayRef || "");
  if (!reference || String(anchor.formula || "") !== slot.publicSnapshot.formula || String(anchor.formulaType || "") !== "dynamicArray") return false;
  if (String(anchor.dynamicArrayRef || "").toUpperCase() !== reference.toUpperCase() || String(anchor.spillRange || "").toUpperCase() !== reference.toUpperCase() || anchor.spillError) return false;
  const bounds = formulaRangeBounds(reference, `${sheet.name}!${address}`);
  const matrix = anchor.spillValues;
  if (!Array.isArray(matrix) || matrix.length !== bounds.bottom - bounds.top + 1 || matrix.some((row) => !Array.isArray(row) || row.length !== bounds.right - bounds.left + 1)) return false;
  const owner = `${String(sheet.name || "").replaceAll("'", "''")}!${address.toUpperCase()}`;
  for (let row = bounds.top; row <= bounds.bottom; row += 1) {
    for (let column = bounds.left; column <= bounds.right; column += 1) {
      const candidateAddress = cellAddress(row, column);
      const candidate = sheet.store.get(candidateAddress);
      if (row === bounds.top && column === bounds.left) {
        if (candidate !== anchor) return false;
        continue;
      }
      if (candidate.spillParent !== owner || candidate.spillAnchor !== address || String(candidate.spillRange || "").toUpperCase() !== reference.toUpperCase()) return false;
      if (candidate.formula) return false;
      if (!Object.is(candidate.value ?? null, matrix[row - bounds.top]?.[column - bounds.left] ?? null)) return false;
    }
  }
  return true;
}

function sourceBoundFormulaCellSnapshot(cell) {
  return {
    value: cell.value instanceof Date ? { type: "date", value: cell.value.getTime() } : cell.value ?? null,
    formula: cell.formula == null ? null : String(cell.formula),
    formulaType: cell.formulaType == null ? null : String(cell.formulaType),
    sharedIndex: cell.sharedIndex == null ? null : Number(cell.sharedIndex),
    sharedRef: cell.sharedRef == null ? null : String(cell.sharedRef),
    arrayRef: cell.arrayRef == null ? null : String(cell.arrayRef),
    dynamicArrayRef: cell.dynamicArrayRef == null ? null : String(cell.dynamicArrayRef),
    spillParent: cell.spillParent == null ? null : String(cell.spillParent),
    spillAnchor: cell.spillAnchor == null ? null : String(cell.spillAnchor),
    spillRange: cell.spillRange == null ? null : String(cell.spillRange),
    spillError: cell.spillError == null ? null : String(cell.spillError),
    style: cell.style || {},
  };
}

function unsupportedWorkbookFeatures(workbook, state) {
  const unsupported = [];
  if (workbook.indexedColors?.length) unsupported.push("custom indexed colors");
  if (workbook.connections?.length && !state) unsupported.push("source-free workbook connections");
  for (const sheet of workbook.worksheets?.items || []) {
    const prefix = `worksheet ${sheet.name}`;
    if (sheet.shapes?.length) unsupported.push(`${prefix} shapes`);
    for (const [address, cell] of sheet.store?.entries?.() || []) {
      if (cell.style && Object.keys(cell.style).some((key) => cell.style[key] != null)) wireCellStyle(cell.style, `${sheet.name}!${address}`);
      const metadata = Object.keys(cell).filter((key) => !["value", "formula", "style"].includes(key) && !XLSX_FORMULA_METADATA_KEYS.has(key));
      if (metadata.length) unsupported.push(`${prefix} advanced formula metadata at ${address}`);
    }
  }
  return unsupported;
}

const XLSX_DATA_VALIDATION_TYPES = new Set(["list", "whole", "decimal", "date", "time", "textLength", "custom"]);
const XLSX_CONDITIONAL_FORMAT_TYPES = new Set(["cellIs", "expression", "containsText", "colorScale", "dataBar", "iconSet"]);
const BRACED_GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/i;

function wireDataValidation(item, sheetName) {
  let rule;
  try {
    rule = normalizeSpreadsheetDataValidationRule(item || {});
  } catch (error) {
    throw new OfficeKitCodecError(`Worksheet ${sheetName} data validation ${item?.id || "(unnamed)"} is invalid: ${error.message}`, [], { code: "invalid_data_validation", cause: error });
  }
  const type = String(rule.type || item?.type || "");
  if (!XLSX_DATA_VALIDATION_TYPES.has(type)) throw new OfficeKitCodecError(`Worksheet ${sheetName} data validation ${item?.id || "(unnamed)"} uses unsupported type ${type || "(empty)"}.`, [], { code: "unsupported_data_validation" });
  const values = Array.isArray(rule.values) ? rule.values.map(String) : [];
  if (values.length && rule.formula1 != null) throw new OfficeKitCodecError(`Worksheet ${sheetName} data validation ${item?.id || "(unnamed)"} cannot combine values and formula1.`, [], { code: "invalid_data_validation" });
  return {
    id: String(item?.id || ""),
    range: String(item?.range || "A1"),
    type,
    operator: String(rule.operator || item?.operator || ""),
    formula1: rule.formula1 == null ? "" : String(rule.formula1),
    formula2: rule.formula2 == null ? "" : String(rule.formula2),
    values,
    ...(Object.hasOwn(rule, "allowBlank") ? { allowBlank: rule.allowBlank } : {}),
    ...(Object.hasOwn(rule, "showInputMessage") ? { showInputMessage: rule.showInputMessage } : {}),
    ...(Object.hasOwn(rule, "promptTitle") ? { promptTitle: rule.promptTitle } : {}),
    ...(Object.hasOwn(rule, "prompt") ? { prompt: rule.prompt } : {}),
    ...(Object.hasOwn(rule, "showErrorMessage") ? { showErrorMessage: rule.showErrorMessage } : {}),
    ...(Object.hasOwn(rule, "errorTitle") ? { errorTitle: rule.errorTitle } : {}),
    ...(Object.hasOwn(rule, "error") ? { error: rule.error } : {}),
    ...(Object.hasOwn(rule, "errorStyle") ? { errorStyle: rule.errorStyle } : {}),
    ...(Object.hasOwn(rule, "showDropdown") ? { showDropdown: rule.showDropdown } : {}),
  };
}

function publicDataValidation(item) {
  return {
    id: item.id || undefined,
    range: item.range || "A1",
    rule: {
      type: item.type,
      ...(item.operator ? { operator: item.operator } : {}),
      ...(item.formula1 ? { formula1: item.formula1 } : {}),
      ...(item.formula2 ? { formula2: item.formula2 } : {}),
      ...(item.values?.length ? { values: [...item.values] } : {}),
      ...(item.allowBlank !== undefined ? { allowBlank: item.allowBlank } : {}),
      ...(item.showInputMessage !== undefined ? { showInputMessage: item.showInputMessage } : {}),
      ...(item.promptTitle ? { promptTitle: item.promptTitle } : {}),
      ...(item.prompt ? { prompt: item.prompt } : {}),
      ...(item.showErrorMessage !== undefined ? { showErrorMessage: item.showErrorMessage } : {}),
      ...(item.errorTitle ? { errorTitle: item.errorTitle } : {}),
      ...(item.error ? { error: item.error } : {}),
      ...(item.errorStyle ? { errorStyle: item.errorStyle } : {}),
      ...(item.showDropdown !== undefined ? { showDropdown: item.showDropdown } : {}),
    },
  };
}

function wireConditionalFormat(item, sheetName, index) {
  const rule = item?.rule || item || {};
  const ruleType = String(item?.ruleType || rule.ruleType || rule.type || "expression");
  if (!XLSX_CONDITIONAL_FORMAT_TYPES.has(ruleType)) throw new OfficeKitCodecError(`Worksheet ${sheetName} conditional format ${item?.id || "(unnamed)"} uses unsupported type ${ruleType}.`, [], { code: "unsupported_conditional_format" });
  const rawFormulas = item?.formulas ?? rule.formulas ?? item?.formula ?? item?.expression ?? rule.formula ?? rule.expression;
  const formulas = (Array.isArray(rawFormulas) ? rawFormulas : rawFormulas == null ? [] : [rawFormulas]).map(String);
  const colors = (item?.colors || rule.colors || []).map((color, colorIndex) => wireSpreadsheetColor(color, `${sheetName}!${item?.range || "A1"}`, `conditional format color ${colorIndex + 1}`));
  let dataBar;
  let iconSet;
  try {
    if (ruleType === "dataBar") {
      const profile = normalizeDataBarConfig(rule, `Worksheet ${sheetName} dataBar ${item?.id || "(unnamed)"}`);
      dataBar = {
        color: wireSpreadsheetColor(profile.color, `${sheetName}!${item?.range || "A1"}`, "data-bar"),
        thresholds: profile.thresholds.map((threshold) => ({ ...threshold })),
        ...(profile.showValue == null ? {} : { showValue: profile.showValue }),
        ...(profile.gradient == null ? {} : { gradient: profile.gradient }),
      };
    }
    if (ruleType === "iconSet") {
      const profile = normalizeIconSetConfig(rule, `Worksheet ${sheetName} iconSet ${item?.id || "(unnamed)"}`);
      iconSet = {
        iconSet: profile.iconSet,
        thresholds: profile.thresholds.map((threshold) => ({ ...threshold })),
        ...(profile.showValue == null ? {} : { showValue: profile.showValue }),
        ...(profile.reverse == null ? {} : { reverse: profile.reverse }),
      };
    }
  } catch (cause) {
    throw new OfficeKitCodecError(cause.message, [], { code: "invalid_conditional_format", cause });
  }
  if (ruleType === "colorScale" && ![2, 3].includes(colors.length)) throw new OfficeKitCodecError(`Worksheet ${sheetName} colorScale ${item?.id || "(unnamed)"} requires two or three colors.`, [], { code: "invalid_conditional_format" });
  if (!["colorScale", "dataBar", "iconSet", "containsText"].includes(ruleType) && formulas.length === 0) throw new OfficeKitCodecError(`Worksheet ${sheetName} conditional format ${item?.id || "(unnamed)"} requires a formula.`, [], { code: "invalid_conditional_format" });
  if (ruleType !== "colorScale" && colors.length) throw new OfficeKitCodecError(`Worksheet ${sheetName} conditional format ${item?.id || "(unnamed)"} colors are valid only for colorScale.`, [], { code: "invalid_conditional_format" });
  if ((dataBar || iconSet) && (formulas.length || item?.format || rule.format || item?.operator || rule.operator || item?.text || rule.text)) throw new OfficeKitCodecError(`Worksheet ${sheetName} ${ruleType} ${item?.id || "(unnamed)"} cannot combine visual metadata with formulas, operators, text, or differential formatting.`, [], { code: "invalid_conditional_format" });
  return {
    id: String(item?.id || ""),
    range: String(item?.range || "A1"),
    ruleType,
    operator: String(item?.operator || rule.operator || ""),
    formulas,
    text: String(item?.text || rule.text || ""),
    format: wireCellStyle(item?.format || rule.format || {}, `${sheetName}!${item?.range || "A1"}`),
    colors,
    priority: Number.isInteger(item?.priority) && item.priority > 0 ? item.priority : index + 1,
    dataBar,
    iconSet,
  };
}

function publicConditionalFormat(item) {
  const formulas = [...(item.formulas || [])];
  const thresholds = (source) => (source?.thresholds || []).map((threshold) => ({ type: threshold.type, ...(threshold.value == null ? {} : { value: threshold.value }) }));
  return {
    id: item.id || undefined,
    range: item.range || "A1",
    ruleType: item.ruleType,
    ...(item.operator ? { operator: item.operator } : {}),
    ...(formulas.length === 1 ? { formula: formulas[0] } : formulas.length ? { formulas } : {}),
    ...(item.text ? { text: item.text } : {}),
    ...(item.format ? { format: cellStyleFromWire(item.format) || {} } : {}),
    ...(item.colors?.length ? { colors: item.colors.map(spreadsheetColorFromWire) } : {}),
    ...(item.dataBar ? {
      color: spreadsheetColorFromWire(item.dataBar.color),
      thresholds: thresholds(item.dataBar),
      ...(item.dataBar.showValue == null ? {} : { showValue: item.dataBar.showValue }),
      ...(item.dataBar.gradient == null ? {} : { gradient: item.dataBar.gradient }),
    } : {}),
    ...(item.iconSet ? {
      iconSet: item.iconSet.iconSet,
      thresholds: thresholds(item.iconSet),
      ...(item.iconSet.showValue == null ? {} : { showValue: item.iconSet.showValue }),
      ...(item.iconSet.reverse == null ? {} : { reverse: item.iconSet.reverse }),
    } : {}),
    ...(item.priority ? { priority: item.priority } : {}),
  };
}

function publicThreadedComment(item) {
  return {
    ...(item.nativeCommentId || item.id ? { id: item.nativeCommentId || item.id } : {}),
    ...(item.personId ? { personId: item.personId } : {}),
    ...(item.dateTime ? { date: item.dateTime } : {}),
    ...(item.parentNativeCommentId ? { parentId: item.parentNativeCommentId } : {}),
    author: item.author || "User",
    person: {
      displayName: item.author || "User",
      ...(item.userId ? { userId: item.userId } : {}),
      ...(item.providerId ? { providerId: item.providerId } : {}),
    },
    done: Boolean(item.resolved),
  };
}

function wireThreadedComments(workbook, sheet) {
  const activeSheetName = workbook.worksheets.getActiveWorksheet().name;
  return (workbook.comments?.threads || []).filter((thread) => (thread.target?.sheetName || activeSheetName) === sheet.name).flatMap((thread) => {
    const address = String(thread.target?.address || "").toUpperCase();
    if (!/^[A-Z]{1,3}[1-9]\d*$/.test(address)) throw new OfficeKitCodecError(`Worksheet ${sheet.name} threaded comment ${thread.id} must target one cell.`, [], { code: "invalid_threaded_comment_target" });
    const comments = thread.comments || [];
    if (!comments.length) throw new OfficeKitCodecError(`Worksheet ${sheet.name} threaded comment ${thread.id} has no root comment.`, [], { code: "invalid_spreadsheet_threaded_comment" });
    const nativeIds = comments.map((comment, index) => BRACED_GUID.test(comment.id || "")
      ? String(comment.id).toUpperCase()
      : deterministicSpreadsheetGuid(`office-kit:${sheet.id}:${thread.id}:${address}:${index}`));
    const rootModelIds = new Set([String(thread.id || ""), String(comments[0]?.id || ""), nativeIds[0]].filter(Boolean));
    return comments.map((comment, index) => {
      if (index > 0 && comment.parentId != null && !rootModelIds.has(String(comment.parentId))) {
        throw new OfficeKitCodecError(`Worksheet ${sheet.name} threaded comment ${thread.id} contains a nested or branched reply graph.`, [], { code: "unsupported_threaded_comment_reply_topology" });
      }
      const person = comment.person || {};
      return {
        id: String(index === 0 ? thread.id || "" : comment.modelId || comment.id || `${thread.id}/reply/${index}`),
        cellReference: address,
        nativeCommentId: nativeIds[index],
        text: String(comment.text ?? ""),
        personId: BRACED_GUID.test(comment.personId || person.id || "") ? String(comment.personId || person.id).toUpperCase() : "",
        author: String(person.displayName || comment.author || thread.author || "User"),
        userId: String(person.userId ?? comment.userId ?? ""),
        providerId: String(person.providerId ?? comment.providerId ?? ""),
        dateTime: comment.date ? new Date(comment.date).toISOString() : "1970-01-01T00:00:00.000Z",
        resolved: Boolean(comment.done ?? thread.resolved),
        ...(index > 0 ? { parentNativeCommentId: nativeIds[0] } : {}),
      };
    });
  });
}

function wireWorkbookTheme(theme, source) {
  let normalized;
  try {
    normalized = normalizeXlsxThemeConfig(theme);
  } catch (cause) {
    throw new OfficeKitCodecError(`Workbook theme is invalid: ${cause.message}`, [], { code: "invalid_workbook_theme", cause });
  }
  return {
    name: normalized.name,
    ...Object.fromEntries(XLSX_THEME_WIRE_FIELDS.map(([model, wire]) => [wire, normalized.colors[model].replace(/^#/, "").toUpperCase()])),
    source,
  };
}

function workbookThemeFromWire(theme) {
  if (!theme || XLSX_THEME_WIRE_FIELDS.some(([, wire]) => !/^[0-9A-Fa-f]{6}$/.test(theme[wire] || ""))) return undefined;
  return normalizeXlsxThemeConfig({
    name: theme.name,
    colors: Object.fromEntries(XLSX_THEME_WIRE_FIELDS.map(([model, wire]) => [model, `#${theme[wire]}`])),
  });
}

function sameWorkbookTheme(left, right) {
  const a = normalizeXlsxThemeConfig(left);
  const b = normalizeXlsxThemeConfig(right);
  return a.name === b.name && XLSX_THEME_COLOR_NAMES.every((name) => a.colors[name] === b.colors[name]);
}

const WORKBOOK_CONNECTION_BOOLEAN_FIELDS = ["keepAlive", "background", "refreshOnLoad", "saveData"];

function publicWorkbookConnection(value) {
  const connection = {
    connectionId: Number(value?.connectionId ?? 0),
    name: String(value?.name ?? ""),
    type: Number(value?.type ?? 0),
    refreshedVersion: Number(value?.refreshedVersion ?? 0),
  };
  if (value?.description !== undefined) connection.description = String(value.description);
  for (const field of WORKBOOK_CONNECTION_BOOLEAN_FIELDS) if (value?.[field] !== undefined) connection[field] = Boolean(value[field]);
  if (value?.intervalMinutes !== undefined) connection.intervalMinutes = Number(value.intervalMinutes);
  return connection;
}

function connectionSnapshot(value) {
  return publicWorkbookConnection(value);
}

function wireSourceBoundConnectionRefreshOnLoad(slot) {
  const source = publicWorkbookConnection(slot.wire);
  if (source.refreshOnLoad !== true) {
    throw new OfficeKitCodecError(
      `Imported workbook connection ${slot.connection.connectionId} may disable refreshOnLoad only when the validated source explicitly enables it.`,
      [],
      { code: "unsupported_workbook_connection_edit" },
    );
  }
  const expected = { ...slot.publicSnapshot, refreshOnLoad: false };
  if (JSON.stringify(connectionSnapshot(slot.connection)) !== JSON.stringify(expected)) {
    throw new OfficeKitCodecError(
      `Imported workbook connection ${slot.connection.connectionId} is source-bound; only refreshOnLoad may change from true to false.`,
      [],
      { code: "unsupported_workbook_connection_edit" },
    );
  }
  return { ...slot.wire, refreshOnLoad: false };
}

function wireWorkbookConnections(workbook, state) {
  const remaining = new Set(workbook.connections || []);
  const output = [];
  for (const slot of state?.connectionSlots || []) {
    if (!remaining.delete(slot.connection)) {
      throw new OfficeKitCodecError(`Workbook cannot remove imported connection ${slot.connection.connectionId} in the bounded OfficeKit slice.`, [], { code: "invalid_workbook_connection" });
    }
    if (JSON.stringify(connectionSnapshot(slot.connection)) === JSON.stringify(slot.publicSnapshot)) output.push(slot.wire);
    else output.push(wireSourceBoundConnectionRefreshOnLoad(slot));
  }
  if (remaining.size) {
    throw new OfficeKitCodecError("OfficeKit cannot author workbook connections; imported connections may only disable explicit refreshOnLoad=true.", [], { code: "unsupported_workbook_connection_edit" });
  }
  return output;
}

function publicWorkbookDefinedName(value) {
  const definedName = {
    id: String(value?.id ?? ""),
    name: String(value?.name ?? ""),
    refersTo: String(value?.refersTo ?? ""),
  };
  if (value?.scopeSheetName !== undefined) definedName.scope = String(value.scopeSheetName);
  if (value?.comment !== undefined) definedName.comment = String(value.comment);
  if (value?.hidden !== undefined) definedName.hidden = Boolean(value.hidden);
  return definedName;
}

function definedNameSnapshot(value) {
  return {
    id: String(value?.id ?? ""),
    name: String(value?.name ?? ""),
    refersTo: String(value?.refersTo ?? ""),
    ...(value?.scope !== undefined ? { scope: String(value.scope) } : {}),
    ...(value?.comment !== undefined ? { comment: String(value.comment) } : {}),
    ...(value?.hidden !== undefined ? { hidden: Boolean(value.hidden) } : {}),
  };
}

function wireWorkbookDefinedName(value, source) {
  const publicValue = definedNameSnapshot(value);
  return {
    id: publicValue.id,
    name: publicValue.name,
    refersTo: publicValue.refersTo,
    ...(publicValue.scope !== undefined ? { scopeSheetName: publicValue.scope } : {}),
    ...(publicValue.comment !== undefined ? { comment: publicValue.comment } : {}),
    ...(publicValue.hidden !== undefined ? { hidden: publicValue.hidden } : {}),
    source,
  };
}

function wireWorkbookDefinedNames(workbook, state) {
  const remaining = new Set(workbook.definedNames?.items || []);
  const output = [];
  for (const slot of state?.definedNameSlots || []) {
    if (!remaining.delete(slot.definedName)) {
      throw new OfficeKitCodecError(`Workbook cannot remove imported defined name ${slot.definedName.name} in the bounded OfficeKit slice.`, [], { code: "invalid_workbook_defined_name" });
    }
    output.push(JSON.stringify(definedNameSnapshot(slot.definedName)) === JSON.stringify(slot.publicSnapshot)
      ? slot.wire
      : wireWorkbookDefinedName(slot.definedName, slot.wire.source));
  }
  output.push(...[...remaining].map((definedName) => wireWorkbookDefinedName(definedName)));
  return output;
}

function publicWorkbookCalculation(value) {
  if (!value) return undefined;
  const calculation = {};
  if (value.mode !== undefined) {
    calculation.mode = value.mode === SpreadsheetCalculationMode.AUTOMATIC ? "automatic"
      : value.mode === SpreadsheetCalculationMode.AUTOMATIC_EXCEPT_TABLES ? "automaticExceptTables"
      : value.mode === SpreadsheetCalculationMode.MANUAL ? "manual" : undefined;
    if (!calculation.mode) throw new OfficeKitCodecError("OfficeKit returned an unsupported workbook calculation mode.", [], { code: "invalid_workbook_calculation" });
  }
  for (const [wireField, publicField] of [["calculateOnSave", "calculateOnSave"], ["fullCalculationOnLoad", "fullCalculationOnLoad"], ["forceFullCalculation", "forceFullCalculation"], ["fullPrecision", "fullPrecision"]])
    if (value[wireField] !== undefined) calculation[publicField] = Boolean(value[wireField]);
  const iteration = {};
  if (value.iterationEnabled !== undefined) iteration.enabled = Boolean(value.iterationEnabled);
  if (value.maxIterations !== undefined) iteration.maxIterations = value.maxIterations;
  if (value.maxChange !== undefined) iteration.maxChange = value.maxChange;
  if (Object.keys(iteration).length) calculation.iteration = iteration;
  return calculation;
}

function calculationSnapshot(value) {
  if (value === undefined) return undefined;
  return {
    ...(value.mode !== undefined ? { mode: value.mode } : {}),
    ...(value.calculateOnSave !== undefined ? { calculateOnSave: Boolean(value.calculateOnSave) } : {}),
    ...(value.fullCalculationOnLoad !== undefined ? { fullCalculationOnLoad: Boolean(value.fullCalculationOnLoad) } : {}),
    ...(value.forceFullCalculation !== undefined ? { forceFullCalculation: Boolean(value.forceFullCalculation) } : {}),
    ...(value.iteration ? { iteration: {
      ...(value.iteration.enabled !== undefined ? { enabled: Boolean(value.iteration.enabled) } : {}),
      ...(value.iteration.maxIterations !== undefined ? { maxIterations: Number(value.iteration.maxIterations) } : {}),
      ...(value.iteration.maxChange !== undefined ? { maxChange: Number(value.iteration.maxChange) } : {}),
    } } : {}),
    ...(value.fullPrecision !== undefined ? { fullPrecision: Boolean(value.fullPrecision) } : {}),
  };
}

function wireWorkbookCalculation(value, source) {
  if (value === undefined) return undefined;
  const calculation = calculationSnapshot(value);
  const mode = calculation.mode === "automatic" ? SpreadsheetCalculationMode.AUTOMATIC
    : calculation.mode === "automaticExceptTables" ? SpreadsheetCalculationMode.AUTOMATIC_EXCEPT_TABLES
    : calculation.mode === "manual" ? SpreadsheetCalculationMode.MANUAL : undefined;
  if (calculation.mode !== undefined && mode === undefined) throw new OfficeKitCodecError(`Unsupported workbook calculation mode ${calculation.mode}.`, [], { code: "invalid_workbook_calculation" });
  return {
    ...(mode !== undefined ? { mode } : {}),
    ...(calculation.calculateOnSave !== undefined ? { calculateOnSave: calculation.calculateOnSave } : {}),
    ...(calculation.fullCalculationOnLoad !== undefined ? { fullCalculationOnLoad: calculation.fullCalculationOnLoad } : {}),
    ...(calculation.forceFullCalculation !== undefined ? { forceFullCalculation: calculation.forceFullCalculation } : {}),
    ...(calculation.iteration?.enabled !== undefined ? { iterationEnabled: calculation.iteration.enabled } : {}),
    ...(calculation.iteration?.maxIterations !== undefined ? { maxIterations: calculation.iteration.maxIterations } : {}),
    ...(calculation.iteration?.maxChange !== undefined ? { maxChange: calculation.iteration.maxChange } : {}),
    ...(calculation.fullPrecision !== undefined ? { fullPrecision: calculation.fullPrecision } : {}),
    source,
  };
}

function wireWorkbookCalculationForExport(workbook, state) {
  const slot = state?.calculationSlot;
  if (!slot) return wireWorkbookCalculation(workbook.calculation);
  if (workbook.calculation === undefined) throw new OfficeKitCodecError("Workbook cannot remove imported calculation properties in the bounded OfficeKit slice.", [], { code: "invalid_workbook_calculation" });
  return JSON.stringify(calculationSnapshot(workbook.calculation)) === JSON.stringify(slot.publicSnapshot)
    ? slot.wire
    : wireWorkbookCalculation(workbook.calculation, slot.wire.source);
}

function tableColumnNames(table) {
  let bounds;
  try {
    bounds = formulaRangeBounds(table.range, table.name || "worksheet table");
  } catch (cause) {
    throw new OfficeKitCodecError(`Worksheet table ${table.name || "(unnamed)"} has an invalid range: ${cause.message}`, [], { code: "invalid_worksheet_table", cause });
  }
  const count = bounds.right - bounds.left + 1;
  if (Array.isArray(table.columnNames)) return table.columnNames.map((value) => String(value));
  const headers = table.showHeaders !== false && table.values?.[0] ? table.values[0] : [];
  return Array.from({ length: count }, (_value, index) => String(headers[index] ?? "").trim() || `Column${index + 1}`);
}

function tableColumnDefinitions(table, names) {
  if (!Array.isArray(table.columnDefinitions)) return undefined;
  return names.map((name, index) => {
    const column = table.columnDefinitions[index] || {};
    return {
      name,
      calculatedColumnFormula: column.calculatedColumnFormula ? String(column.calculatedColumnFormula) : "",
      calculatedColumnFormulaArray: Boolean(column.calculatedColumnFormulaArray),
      totalsRowFunction: column.totalsRowFunction ? String(column.totalsRowFunction) : "",
      totalsRowLabel: column.totalsRowLabel ? String(column.totalsRowLabel) : "",
      totalsRowFormula: column.totalsRowFormula ? String(column.totalsRowFormula) : "",
      totalsRowFormulaArray: Boolean(column.totalsRowFormulaArray),
    };
  });
}

function tableFilters(table) {
  if (!Array.isArray(table.filters)) return [];
  return table.filters.map((filter) => {
    const columnIndex = Number(filter?.columnIndex ?? 0);
    if (filter?.kind === "custom") {
      return {
        columnIndex,
        criteria: {
          case: "custom",
          value: {
            matchAll: Boolean(filter.matchAll),
            criteria: Array.isArray(filter.criteria)
              ? filter.criteria.map((criterion) => ({ operator: String(criterion?.operator ?? ""), value: String(criterion?.value ?? "") }))
              : [],
          },
        },
      };
    }
    if (filter?.kind === "dynamic") {
      return {
        columnIndex,
        criteria: {
          case: "dynamic",
          value: {
            type: String(filter.type ?? ""),
            value: filter.value == null ? undefined : Number(filter.value),
            maxValue: filter.maxValue == null ? undefined : Number(filter.maxValue),
          },
        },
      };
    }
    if (filter?.kind === "top10") {
      return {
        columnIndex,
        criteria: {
          case: "top10",
          value: {
            top: filter.top ?? true,
            percent: Boolean(filter.percent),
            value: Number(filter.value ?? 0),
            filterValue: filter.filterValue == null ? undefined : Number(filter.filterValue),
          },
        },
      };
    }
    if (filter?.kind === "icon") {
      return {
        columnIndex,
        criteria: {
          case: "icon",
          value: {
            iconSet: String(filter.iconSet ?? ""),
            iconId: filter.iconId == null ? undefined : Number(filter.iconId),
          },
        },
      };
    }
    if (filter?.kind === "color") {
      return {
        columnIndex,
        criteria: { case: "color", value: wireTableColor(filter, `table ${table.name} filter column ${columnIndex}`) },
      };
    }
    return {
      columnIndex,
      criteria: {
        case: "values",
        value: {
          values: Array.isArray(filter?.values) ? filter.values.map((value) => String(value)) : [],
          includeBlank: Boolean(filter?.includeBlank),
          dateGroups: Array.isArray(filter?.dateGroups) ? filter.dateGroups.map((group) => ({
            grouping: String(group?.grouping ?? ""),
            year: Number(group?.year ?? 0),
            month: group?.month == null ? undefined : Number(group.month),
            day: group?.day == null ? undefined : Number(group.day),
            hour: group?.hour == null ? undefined : Number(group.hour),
            minute: group?.minute == null ? undefined : Number(group.minute),
            second: group?.second == null ? undefined : Number(group.second),
          })) : [],
          calendarType: filter?.calendarType ? String(filter.calendarType) : "",
        },
      },
    };
  });
}

function publicTableFilter(filter) {
  if (filter?.criteria?.case === "custom") {
    return {
      columnIndex: Number(filter.columnIndex ?? 0),
      kind: "custom",
      matchAll: Boolean(filter.criteria.value?.matchAll),
      criteria: (filter.criteria.value?.criteria || []).map((criterion) => ({ operator: criterion.operator, value: criterion.value })),
    };
  }
  if (filter?.criteria?.case === "dynamic") {
    return {
      columnIndex: Number(filter.columnIndex ?? 0),
      kind: "dynamic",
      type: filter.criteria.value?.type || "",
      ...(filter.criteria.value?.value == null ? {} : { value: filter.criteria.value.value }),
      ...(filter.criteria.value?.maxValue == null ? {} : { maxValue: filter.criteria.value.maxValue }),
    };
  }
  if (filter?.criteria?.case === "top10") {
    return {
      columnIndex: Number(filter.columnIndex ?? 0),
      kind: "top10",
      top: Boolean(filter.criteria.value?.top),
      percent: Boolean(filter.criteria.value?.percent),
      value: Number(filter.criteria.value?.value ?? 0),
      ...(filter.criteria.value?.filterValue == null ? {} : { filterValue: filter.criteria.value.filterValue }),
    };
  }
  if (filter?.criteria?.case === "icon") {
    return {
      columnIndex: Number(filter.columnIndex ?? 0),
      kind: "icon",
      iconSet: filter.criteria.value?.iconSet || "",
      ...(filter.criteria.value?.iconId == null ? {} : { iconId: filter.criteria.value.iconId }),
    };
  }
  if (filter?.criteria?.case === "color") {
    return {
      columnIndex: Number(filter.columnIndex ?? 0),
      kind: "color",
      ...publicTableColor(filter.criteria.value),
    };
  }
  return {
    columnIndex: Number(filter?.columnIndex ?? 0),
    kind: "values",
    values: [...(filter?.criteria?.value?.values || [])],
    includeBlank: Boolean(filter?.criteria?.value?.includeBlank),
    ...((filter?.criteria?.value?.dateGroups || []).length ? {
      dateGroups: filter.criteria.value.dateGroups.map((group) => ({
        grouping: group.grouping,
        year: Number(group.year ?? 0),
        ...(group.month == null ? {} : { month: group.month }),
        ...(group.day == null ? {} : { day: group.day }),
        ...(group.hour == null ? {} : { hour: group.hour }),
        ...(group.minute == null ? {} : { minute: group.minute }),
        ...(group.second == null ? {} : { second: group.second }),
      })),
    } : {}),
    ...(filter?.criteria?.value?.calendarType ? { calendarType: filter.criteria.value.calendarType } : {}),
  };
}

function wireTableSortState(sort, address) {
  if (!sort) return undefined;
  return {
    reference: String(sort.reference ?? ""),
    caseSensitive: Boolean(sort.caseSensitive),
    ...(sort.sortMethod == null ? {} : { sortMethod: String(sort.sortMethod) }),
    ...(sort.columnSort == null ? {} : { columnSort: Boolean(sort.columnSort) }),
    conditions: Array.isArray(sort.conditions)
      ? sort.conditions.map((condition) => ({
          reference: String(condition?.reference ?? ""),
          descending: Boolean(condition?.descending),
          ...((condition?.kind === "icon" || condition?.iconSet) ? {
            icon: {
              iconSet: String(condition.iconSet ?? ""),
              iconId: condition.iconId == null ? undefined : Number(condition.iconId),
            },
          } : condition?.kind === "color" ? {
            color: wireTableColor(condition, `${address} sort ${condition.reference}`),
          } : condition?.customList == null ? {} : { customList: String(condition.customList) }),
        }))
      : [],
  };
}

function tableSortState(table) {
  return wireTableSortState(table?.sortState, `table ${table?.name || "(unnamed)"}`);
}

function publicTableSortState(sort) {
  if (!sort) return undefined;
  return {
    reference: sort.reference,
    caseSensitive: Boolean(sort.caseSensitive),
    ...(sort.sortMethod == null ? {} : { sortMethod: sort.sortMethod }),
    ...(sort.columnSort == null ? {} : { columnSort: Boolean(sort.columnSort) }),
    conditions: (sort.conditions || []).map((condition) => ({
      reference: condition.reference,
      descending: Boolean(condition.descending),
      ...(condition.icon ? {
        kind: "icon",
        iconSet: condition.icon.iconSet,
        ...(condition.icon.iconId == null ? {} : { iconId: condition.icon.iconId }),
      } : condition.color ? {
        kind: "color",
        ...publicTableColor(condition.color),
      } : condition.customList == null ? {} : { customList: condition.customList }),
    })),
  };
}

function wireTableColor(value, address) {
  const target = value?.target;
  if (target !== "cell" && target !== "font") {
    throw new OfficeKitCodecError(`Worksheet ${address} color target must be 'cell' or 'font'.`, [], { code: "invalid_worksheet_table" });
  }
  const color = wireSpreadsheetColor(value.color, address, `${target} color`);
  if (!color) throw new OfficeKitCodecError(`Worksheet ${address} must provide a color.`, [], { code: "invalid_worksheet_table" });
  return { target: { case: target === "cell" ? "cellColor" : "fontColor", value: true }, color };
}

function publicTableColor(value) {
  return {
    target: value?.target?.case === "cellColor" ? "cell" : "font",
    color: spreadsheetColorFromWire(value?.color),
  };
}

const TABLE_QUERY_BOOLEAN_FIELDS = [
  "headers", "rowNumbers", "disableRefresh", "backgroundRefresh", "firstBackgroundRefresh", "refreshOnLoad",
  "fillFormulas", "removeDataOnSave", "disableEdit", "preserveFormatting", "adjustColumnWidth", "intermediate",
  "applyNumberFormats", "applyBorderFormats", "applyFontFormats", "applyPatternFormats", "applyAlignmentFormats",
  "applyWidthHeightFormats",
];
const TABLE_QUERY_REFRESH_POLICY_VALUES = Object.freeze({
  disableRefresh: true,
  backgroundRefresh: false,
  firstBackgroundRefresh: false,
  refreshOnLoad: false,
});

const TABLE_QUERY_REFRESH_BOOLEAN_FIELDS = ["preserveSortFilterLayout", "fieldIdWrapped", "headersInLastRefresh"];
const TABLE_QUERY_REFRESH_UINT_FIELDS = ["minimumVersion", "nextId", "unboundColumnsLeft", "unboundColumnsRight"];
const TABLE_QUERY_FIELD_BOOLEAN_FIELDS = ["dataBound", "rowNumbers", "fillFormulas", "clipped"];

function publicTableQueryField(value) {
  const field = { id: Number(value?.id ?? 0) };
  if (value?.name !== undefined) field.name = String(value.name);
  for (const name of TABLE_QUERY_FIELD_BOOLEAN_FIELDS) if (value?.[name] !== undefined) field[name] = Boolean(value[name]);
  if (value?.tableColumnId !== undefined) field.tableColumnId = Number(value.tableColumnId);
  return field;
}

function publicTableQueryRefresh(value) {
  if (!value) return undefined;
  const refresh = { fields: Array.isArray(value.fields) ? value.fields.map(publicTableQueryField) : [] };
  for (const field of TABLE_QUERY_REFRESH_BOOLEAN_FIELDS) if (value[field] !== undefined) refresh[field] = Boolean(value[field]);
  for (const field of TABLE_QUERY_REFRESH_UINT_FIELDS) if (value[field] !== undefined) refresh[field] = Number(value[field]);
  if (Array.isArray(value.deletedFieldNames) && value.deletedFieldNames.length)
    refresh.deletedFieldNames = value.deletedFieldNames.map((name) => String(name));
  if (value.sortState) refresh.sortState = publicTableSortState(value.sortState);
  return refresh;
}

function publicTableQuery(value) {
  if (!value) return undefined;
  const query = { name: String(value.name ?? ""), connectionId: Number(value.connectionId ?? 0) };
  for (const field of TABLE_QUERY_BOOLEAN_FIELDS) if (value[field] !== undefined) query[field] = Boolean(value[field]);
  if (value.growShrinkType !== undefined) query.growShrinkType = String(value.growShrinkType);
  if (value.autoFormatId !== undefined) query.autoFormatId = Number(value.autoFormatId);
  if (value.refresh) query.refresh = publicTableQueryRefresh(value.refresh);
  return query;
}

function wireSourceBoundQueryRefreshPolicy(table, wire, query, snapshot) {
  const source = publicTableQuery(wire);
  if (!source || !query) {
    throw new OfficeKitCodecError(`Imported query table ${table.name} is source-bound and cannot change its topology.`, [], { code: "unsupported_query_table_edit" });
  }
  const changes = [];
  const normalized = { ...query };
  for (const [field, requiredValue] of Object.entries(TABLE_QUERY_REFRESH_POLICY_VALUES)) {
    const sourceHas = Object.hasOwn(source, field);
    const nextHas = Object.hasOwn(query, field);
    if (sourceHas !== nextHas || source[field] !== query[field]) {
      if (!nextHas || query[field] !== requiredValue) {
        throw new OfficeKitCodecError(`Imported query table ${table.name} may only harden ${field} to ${requiredValue}.`, [], { code: "unsupported_query_table_edit" });
      }
      changes.push(field);
    }
    if (sourceHas) normalized[field] = source[field];
    else delete normalized[field];
  }
  if (!changes.length || JSON.stringify(normalized) !== JSON.stringify(snapshot)) {
    throw new OfficeKitCodecError(`Imported query table ${table.name} is source-bound: only explicit refresh-policy hardening is supported.`, [], { code: "unsupported_query_table_edit" });
  }
  const output = { ...wire };
  for (const field of changes) output[field] = query[field];
  return output;
}

function wireTableQuery(table) {
  const state = table[TABLE_STATE];
  const query = publicTableQuery(table.queryTable);
  if (state) {
    if (JSON.stringify(query) === JSON.stringify(state.querySnapshot)) return state.wire?.queryTable;
    return wireSourceBoundQueryRefreshPolicy(table, state.wire?.queryTable, query, state.querySnapshot);
  }
  if (query) {
    throw new OfficeKitCodecError(`OfficeKit cannot author query table ${table.name}; only recognized imported query tables may apply refresh-policy hardening.`, [], { code: "unsupported_query_table_edit" });
  }
  return undefined;
}

function tableSnapshot(table) {
  const columnNames = tableColumnNames(table);
  return {
    id: table.id,
    name: table.name,
    reference: table.range,
    hasHeaders: table.showHeaders !== false,
    showTotals: Boolean(table.showTotals),
    showFilterButton: table.showFilterButton !== false,
    styleName: table.style || "TableStyleMedium2",
    showFirstColumn: Boolean(table.showFirstColumn),
    showLastColumn: Boolean(table.showLastColumn),
    showRowStripes: table.showRowStripes ?? table.showHeaders !== false,
    showColumnStripes: Boolean(table.showBandedColumns),
    columnNames,
    columns: tableColumnDefinitions(table, columnNames),
    filters: tableFilters(table),
    sortState: tableSortState(table),
    queryTable: publicTableQuery(table.queryTable),
  };
}

function sameTableSnapshot(table, snapshot) {
  return JSON.stringify(tableSnapshot(table)) === JSON.stringify(snapshot);
}

function wireWorksheetTable(table) {
  return { ...tableSnapshot(table), queryTable: wireTableQuery(table), source: table[TABLE_STATE]?.wire?.source };
}

function wireWorksheetTables(sheet, state) {
  const remaining = new Set(sheet.tables?.items || []);
  const output = [];
  for (const slot of state?.slots || []) {
    if (!slot.table) {
      output.push(slot.wire);
      continue;
    }
    if (!remaining.delete(slot.table)) {
      throw new OfficeKitCodecError(`Worksheet ${sheet.name} cannot remove imported table ${slot.table.name} in the bounded OfficeKit slice.`, [], { code: "invalid_worksheet_table" });
    }
    output.push(sameTableSnapshot(slot.table, slot.publicSnapshot) ? slot.wire : wireWorksheetTable(slot.table));
  }
  output.push(...[...remaining].map(wireWorksheetTable));
  return output;
}

function excelSerialFromDate(value, dateSystem, address) {
  const milliseconds = value.getTime();
  if (!Number.isFinite(milliseconds)) throw new OfficeKitCodecError(`Cell ${address} has an invalid Date value.`, [], { code: "invalid_cell_date" });
  const dayMilliseconds = 86_400_000;
  if (dateSystem === "1904") return (milliseconds - Date.UTC(1904, 0, 1)) / dayMilliseconds;
  const serial = (milliseconds - Date.UTC(1899, 11, 31)) / dayMilliseconds;
  return milliseconds >= Date.UTC(1900, 2, 1) ? serial + 1 : serial;
}

function wireCell(address, cell, dateSystem) {
  const coordinates = cellCoordinates(address);
  const dateValue = cell.value instanceof Date;
  const target = {
    row: coordinates.row,
    column: coordinates.column,
    formula: cell.formula ? String(cell.formula) : "",
    formulaMetadata: cellFormulaMetadata(address, cell),
    numberFormatCode: cellNumberFormatCode(cell, address) || (dateValue ? "yyyy-mm-dd hh:mm:ss" : ""),
    style: wireCellStyle(cell.style, address),
    value: { case: undefined },
  };
  if (cell.value == null) return target;
  if (dateValue) {
    target.value = { case: "numberValue", value: excelSerialFromDate(cell.value, dateSystem, address) };
  } else if (typeof cell.value === "string") {
    target.value = EXCEL_ERRORS.has(cell.value) ? { case: "errorValue", value: cell.value } : { case: "stringValue", value: cell.value };
  } else if (typeof cell.value === "number") {
    if (!Number.isFinite(cell.value)) throw new OfficeKitCodecError(`Cell ${address} has a non-finite numeric value.`, [], { code: "non_finite_cell_value" });
    target.value = { case: "numberValue", value: cell.value };
  } else if (typeof cell.value === "boolean") {
    target.value = { case: "boolValue", value: cell.value };
  } else {
    throw new OfficeKitCodecError(`Cell ${address} has unsupported ${cell.value?.constructor?.name || typeof cell.value} content.`, [], { code: "unsupported_cell_value" });
  }
  return target;
}

function workbookEnvelope(workbook) {
  if (!(workbook instanceof Workbook)) throw new TypeError("exportXlsxWithOfficeKit expects a Workbook instance.");
  if (!workbook.worksheets?.items?.length) throw new OfficeKitCodecError("Workbook must contain at least one worksheet.", [], { code: "missing_worksheets" });
  if (!workbook.worksheets.items.some((sheet) => sheet.visibility === "visible")) throw new OfficeKitCodecError("Workbook must contain at least one visible worksheet.", [], { code: "missing_visible_worksheet" });
  const state = workbook[WORKBOOK_STATE];
  assertTrustedImportedState(state, "XLSX");
  const unsupported = unsupportedWorkbookFeatures(workbook, state);
  if (unsupported.length) {
    throw new OfficeKitCodecError(`OfficeKit cannot encode these XLSX features: ${unsupported.slice(0, 8).join(", ")}${unsupported.length > 8 ? `, and ${unsupported.length - 8} more` : ""}. This operation fails closed; preserve them only through a validated source-bound package.`, [], { code: "unsupported_workbook_features" });
  }
  const theme = state?.themeWire && sameWorkbookTheme(workbook.theme, state.publicTheme)
    ? state.themeWire
    : wireWorkbookTheme(workbook.theme, state?.themeWire?.source);
  const views = wireWorkbookViews(workbook, state);
  const assets = new Map();
  const worksheets = workbook.worksheets.items.map((sheet) => {
    const worksheetSlot = state?.worksheetSlots?.get(sheet.id);
    const metadata = wireWorksheetMetadata(sheet, worksheetSlot);
    const cells = (() => {
      const dynamicSlots = state?.dynamicArraySlotsBySheet?.get(sheet.id) || new Map();
      const sourceBoundFormulaSlots = state?.sourceBoundFormulaSlotsBySheet?.get(sheet.id) || new Map();
      const entries = sheet.store?.entries?.() || [];
      const byAddress = new Map(entries);
      const recalculatedDynamicAddresses = new Set();
      for (const [address, slot] of dynamicSlots) {
        const unchanged = byAddress.get(address) === slot.cell
          && JSON.stringify(dynamicArrayCellSnapshot(slot.cell)) === JSON.stringify(slot.publicSnapshot)
          && JSON.stringify(dynamicArrayRangeSnapshot(sheet, address, slot.publicSnapshot.dynamicArrayRef)) === JSON.stringify(slot.publicRangeSnapshot);
        const recalculated = byAddress.get(address) === slot.cell && dynamicArrayRecalculationMatches(sheet, address, slot);
        if (!unchanged && !recalculated) {
          throw new OfficeKitCodecError(`Imported dynamic array ${sheet.name}!${address} is source-bound and read-only in OfficeKit 0.2.`, [], { code: "unsupported_dynamic_array_edit" });
        }
        if (recalculated) recalculatedDynamicAddresses.add(address);
      }
      for (const [address, slot] of sourceBoundFormulaSlots) {
        if (byAddress.get(address) !== slot.cell || JSON.stringify(sourceBoundFormulaCellSnapshot(slot.cell)) !== JSON.stringify(slot.publicSnapshot)) {
          throw new OfficeKitCodecError(`Imported partial shared formula ${sheet.name}!${address} is source-bound and read-only in OfficeKit 0.2.`, [], { code: "unsupported_cell_formula_edit" });
        }
      }
      const output = entries
        .filter(([, cell]) => cell.value != null || cell.formula || cell.formulaType || Object.keys(cell.style || {}).some((key) => cell.style[key] != null))
        .map(([address, cell]) => sourceBoundFormulaSlots.get(address)?.wire || (dynamicSlots.has(address) && !recalculatedDynamicAddresses.has(address) ? dynamicSlots.get(address).wire : wireCell(address, cell, workbook.dateSystem)));
      wireWorksheetDataTables(sheet, state?.dataTablesBySheet?.get(sheet.id), output);
      validateFormulaTopology(output, sheet.name);
      return output;
    })();
    const pivotTables = wireWorksheetPivots(workbook, sheet, state?.pivotsBySheet?.get(sheet.id), cells);
    return {
      id: sheet.id,
      name: sheet.name,
      visibility: metadata.visibility,
      source: metadata.source,
      showGridLines: sheet.showGridLines !== false,
      freezePane: wireWorksheetFreezePane(sheet, worksheetSlot),
      columnDimensions: [...(sheet.columnDimensions || new Map())].map(([column, dimension]) => ({ column, width: dimension.width || 0, hidden: Boolean(dimension.hidden), bestFit: Boolean(dimension.bestFit) })),
      rowDimensions: [...(sheet.rowDimensions || new Map())].map(([row, dimension]) => ({ row, height: dimension.height || 0, hidden: Boolean(dimension.hidden) })),
      mergedRanges: [...(sheet.mergedRanges || [])],
      sortState: wireTableSortState(sheet.sortState, `worksheet ${sheet.name}`),
      tables: wireWorksheetTables(sheet, state?.tablesBySheet?.get(sheet.id)),
      images: wireWorksheetImages(sheet, state?.imagesBySheet?.get(sheet.id), assets),
      charts: wireWorksheetCharts(sheet, state?.chartsBySheet?.get(sheet.id)),
      sparklineGroups: wireWorksheetSparklines(sheet, state?.sparklinesBySheet?.get(sheet.id)),
      pivotTables,
      protection: wireWorksheetProtection(sheet, state?.worksheetSlots?.get(sheet.id)),
      dataValidations: (sheet.dataValidations?.items || []).map((item) => wireDataValidation(item, sheet.name)),
      conditionalFormats: (sheet.conditionalFormattings?.items || []).map((item, index) => wireConditionalFormat(item, sheet.name, index)),
      threadedComments: wireThreadedComments(workbook, sheet),
      cells,
    };
  });
  return {
    protocolVersion: OFFICE_KIT_PROTOCOL_VERSION,
    family: ArtifactFamily.WORKBOOK,
    source: state?.source,
    opaqueOpc: state?.opaqueOpc,
    assets: [...assets.values()],
    diagnostics: state?.diagnostics || [],
    payload: {
      case: "workbook",
      value: {
        id: workbook.id,
        dateSystem: workbook.dateSystem === "1904" ? WorkbookDateSystem.WORKBOOK_DATE_SYSTEM_1904 : WorkbookDateSystem.WORKBOOK_DATE_SYSTEM_1900,
        theme,
        connections: wireWorkbookConnections(workbook, state),
        definedNames: wireWorkbookDefinedNames(workbook, state),
        calculation: wireWorkbookCalculationForExport(workbook, state),
        view: views.view,
        additionalViews: views.additionalViews,
        worksheets,
      },
    },
  };
}

export async function exportXlsxWithOfficeKit(workbook, options = {}) {
  assertCodecOptions(options, new Set(["limits", "recalculate"]), "exportXlsxWithOfficeKit");
  if (!(workbook instanceof Workbook)) throw new TypeError("exportXlsxWithOfficeKit expects a Workbook instance.");
  if (options.recalculate !== false) workbook.recalculate();
  const response = await invokeOfficeKit({
    protocolVersion: OFFICE_KIT_PROTOCOL_VERSION,
    operation: CodecOperation.EXPORT_XLSX,
    family: ArtifactFamily.WORKBOOK,
    artifact: workbookEnvelope(workbook),
    limits: codecLimits(options.limits),
  });
  return new FileBlob(response.file, {
    type: XLSX_MIME,
    metadata: { artifactKind: "workbook", codec: "office-kit", diagnostics: response.diagnostics },
  });
}

function workbookFromEnvelope(envelope) {
  if (envelope.family !== ArtifactFamily.WORKBOOK || envelope.payload.case !== "workbook") {
    throw new OfficeKitCodecError("OfficeKit response does not contain a workbook artifact.", [], { code: "invalid_workbook_artifact" });
  }
  const source = envelope.payload.value;
  const importedTheme = workbookThemeFromWire(source.theme);
  const importedConnections = (source.connections || []).map(publicWorkbookConnection);
  const importedCalculation = publicWorkbookCalculation(source.calculation);
  const workbook = Workbook.create({
    dateSystem: source.dateSystem === WorkbookDateSystem.WORKBOOK_DATE_SYSTEM_1904 ? "1904" : "1900",
    ...(importedTheme ? { theme: importedTheme } : {}),
    connections: importedConnections,
    ...(importedCalculation !== undefined ? { calculation: importedCalculation } : {}),
  });
  workbook.id = source.id || workbook.id;
  const tablesBySheet = new Map();
  const imagesBySheet = new Map();
  const chartsBySheet = new Map();
  const sparklinesBySheet = new Map();
  let pivotsBySheet;
  const dataTablesBySheet = new Map();
  const dynamicArraySlotsBySheet = new Map();
  const sourceBoundFormulaSlotsBySheet = new Map();
  const worksheetSlots = new Map();
  const partialSharedFormulaRangesBySheetName = partialSharedFormulaRanges(envelope.diagnostics);
  const assets = new Map((envelope.assets || []).map((asset) => [asset.id, asset]));
  const connectionSlots = (source.connections || []).map((wire, index) => ({
    wire,
    connection: workbook.connections[index],
    publicSnapshot: connectionSnapshot(workbook.connections[index]),
  }));
  for (const sourceSheet of source.worksheets) {
    const sheet = workbook.worksheets.add(sourceSheet.name, { visibility: publicWorksheetVisibility(sourceSheet.visibility) });
    sheet.id = sourceSheet.id || sheet.id;
    sheet.protection = publicWorksheetProtectionFromWire(sourceSheet.protection);
    worksheetSlots.set(sheet.id, {
      wire: sourceSheet,
      publicSnapshot: worksheetMetadataSnapshot(sheet),
      publicProtectionSnapshot: worksheetProtectionPublicSnapshot(sheet),
    });
    sheet.showGridLines = sourceSheet.showGridLines;
    if (sourceSheet.freezePane) {
      sheet.freezePanes.freezeRows(sourceSheet.freezePane.rows);
      sheet.freezePanes.freezeColumns(sourceSheet.freezePane.columns);
    }
    for (const dimension of sourceSheet.columnDimensions) sheet.columnDimensions.set(dimension.column, { width: dimension.width || undefined, hidden: dimension.hidden, bestFit: dimension.bestFit });
    for (const dimension of sourceSheet.rowDimensions) sheet.rowDimensions.set(dimension.row, { height: dimension.height || undefined, hidden: dimension.hidden });
    sheet.mergedRanges = [...sourceSheet.mergedRanges];
    sheet.sortState = publicTableSortState(sourceSheet.sortState);
    sheet.dataValidations.items = (sourceSheet.dataValidations || []).map(publicDataValidation);
    sheet.conditionalFormattings.items = (sourceSheet.conditionalFormats || []).map(publicConditionalFormat);
    const sourceComments = sourceSheet.threadedComments || [];
    const rootComments = sourceComments.filter((item) => !item.parentNativeCommentId);
    const consumedReplies = new Set();
    for (const sourceComment of rootComments) {
      const thread = workbook.comments.addThread(
        { sheetName: sheet.name, address: sourceComment.cellReference },
        sourceComment.text,
        {
          id: sourceComment.id || undefined,
          author: sourceComment.author || "User",
          resolved: sourceComment.resolved,
          comment: publicThreadedComment(sourceComment),
        },
      );
      if (sourceComment.id) thread.id = sourceComment.id;
      for (const reply of sourceComments.filter((item) => item.parentNativeCommentId === sourceComment.nativeCommentId)) {
        thread.addReply(reply.text, publicThreadedComment(reply));
        consumedReplies.add(reply);
      }
    }
    if (rootComments.length + consumedReplies.size !== sourceComments.length) {
      throw new OfficeKitCodecError(`Worksheet ${sheet.name} contains an unsupported threaded-comment reply graph.`, [], { code: "unsupported_threaded_comment_reply_topology" });
    }
    const dynamicArraySlots = new Map();
    const sourceBoundFormulaSlots = new Map();
    const partialSharedFormulaRanges = partialSharedFormulaRangesBySheetName.get(sourceSheet.name) || [];
    const dataTableSlots = [];
    for (const sourceCell of sourceSheet.cells) {
      const address = cellAddress(sourceCell.row, sourceCell.column);
      const cell = sheet.store.get(address);
      cell.formula = sourceCell.formula || null;
      if (sourceCell.formulaMetadata?.kind === CellFormulaKind.SHARED) {
        cell.formulaType = "shared";
        cell.sharedIndex = sourceCell.formulaMetadata.sharedIndex;
        cell.sharedRef = sourceCell.formulaMetadata.reference;
      } else if (sourceCell.formulaMetadata?.kind === CellFormulaKind.ARRAY) {
        cell.formulaType = "array";
        cell.arrayRef = sourceCell.formulaMetadata.reference;
      } else if (sourceCell.formulaMetadata?.kind === CellFormulaKind.DYNAMIC_ARRAY) {
        cell.formulaType = "dynamicArray";
        cell.dynamicArrayRef = sourceCell.formulaMetadata.reference;
      }
      const staticStyle = cellStyleFromWire(sourceCell.style);
      if (staticStyle || sourceCell.numberFormatCode) cell.style = { ...(staticStyle || {}), ...(sourceCell.numberFormatCode ? { numberFormat: sourceCell.numberFormatCode } : {}) };
      switch (sourceCell.value.case) {
        case "stringValue": cell.value = sourceCell.value.value; break;
        case "numberValue": cell.value = sourceCell.value.value; break;
        case "boolValue": cell.value = sourceCell.value.value; break;
        case "errorValue": cell.value = sourceCell.value.value; break;
        default: cell.value = null;
      }
      if (partialSharedFormulaRanges.some(({ bounds }) =>
        sourceCell.row >= bounds.top && sourceCell.row <= bounds.bottom &&
        sourceCell.column >= bounds.left && sourceCell.column <= bounds.right)) {
        sourceBoundFormulaSlots.set(address, { wire: sourceCell, cell, publicSnapshot: sourceBoundFormulaCellSnapshot(cell) });
      }
      if (sourceCell.formulaMetadata?.kind === CellFormulaKind.DYNAMIC_ARRAY) {
        dynamicArraySlots.set(address, { wire: sourceCell, cell, publicSnapshot: dynamicArrayCellSnapshot(cell) });
      } else if (sourceCell.formulaMetadata?.kind === CellFormulaKind.DATA_TABLE) {
        dataTableSlots.push(hydrateWorksheetDataTable(sheet, sourceCell));
      }
    }
    for (const [address, slot] of dynamicArraySlots) {
      slot.publicRangeSnapshot = dynamicArrayRangeSnapshot(sheet, address, slot.publicSnapshot.dynamicArrayRef);
    }
    dynamicArraySlotsBySheet.set(sheet.id, dynamicArraySlots);
    if (partialSharedFormulaRanges.length && sourceBoundFormulaSlots.size === 0) {
      throw new OfficeKitCodecError(`OfficeKit reported partial shared formulas for ${sourceSheet.name} but returned no matching cells.`, [], { code: "invalid_office_kit_diagnostic" });
    }
    sourceBoundFormulaSlotsBySheet.set(sheet.id, sourceBoundFormulaSlots);
    dataTablesBySheet.set(sheet.id, { slots: dataTableSlots });
    const slots = [];
    for (const sourceTable of sourceSheet.tables || []) {
      if (!sourceTable.name || !sourceTable.reference || !sourceTable.columnNames?.length) {
        slots.push({ wire: sourceTable });
        continue;
      }
      const table = sheet.tables.add({
        id: sourceTable.id,
        range: sourceTable.reference,
        name: sourceTable.name,
        hasHeaders: sourceTable.hasHeaders,
        showTotals: sourceTable.showTotals,
        showFilterButton: sourceTable.showFilterButton,
        showBandedColumns: sourceTable.showColumnStripes,
        style: sourceTable.styleName,
        columnNames: [...sourceTable.columnNames],
        columnDefinitions: sourceTable.columns?.length ? sourceTable.columns.map((column) => ({ ...column })) : undefined,
        filters: sourceTable.filters?.map(publicTableFilter),
        sortState: publicTableSortState(sourceTable.sortState),
        queryTable: publicTableQuery(sourceTable.queryTable),
      });
      table.showHeaders = sourceTable.hasHeaders;
      table.showFirstColumn = sourceTable.showFirstColumn;
      table.showLastColumn = sourceTable.showLastColumn;
      table.showRowStripes = sourceTable.showRowStripes;
      const publicSnapshot = tableSnapshot(table);
      Object.defineProperty(table, TABLE_STATE, { configurable: true, value: { wire: sourceTable, querySnapshot: publicTableQuery(table.queryTable) }, writable: true });
      slots.push({ wire: sourceTable, table, publicSnapshot });
    }
    tablesBySheet.set(sheet.id, { slots });
    const imageSlots = [];
    for (const sourceImage of sourceSheet.images || []) {
      const image = spreadsheetImageFromWire(sheet, sourceImage, assets);
      imageSlots.push({ wire: sourceImage, image, publicSnapshot: spreadsheetImageSnapshot(image) });
    }
    imagesBySheet.set(sheet.id, { slots: imageSlots });
    const chartSlots = [];
    for (const sourceChart of sourceSheet.charts || []) {
      const chart = spreadsheetChartFromWire(sheet, sourceChart);
      chartSlots.push({ wire: sourceChart, chart, publicSnapshot: spreadsheetChartSnapshot(chart) });
    }
    chartsBySheet.set(sheet.id, { slots: chartSlots });
    const sparklineSlots = [];
    for (const sourceSparkline of sourceSheet.sparklineGroups || []) {
      const group = spreadsheetSparklineFromWire(sheet, sourceSparkline);
      sparklineSlots.push({ wire: sourceSparkline, group, publicSnapshot: spreadsheetSparklineSnapshot(group) });
    }
    sparklinesBySheet.set(sheet.id, { slots: sparklineSlots });
  }
  for (const sourceDefinedName of source.definedNames || []) workbook.definedNames.add(publicWorkbookDefinedName(sourceDefinedName));
  pivotsBySheet = hydrateWorkbookPivots(workbook, source.worksheets);
  const sourceViews = source.view ? [source.view, ...(source.additionalViews || [])] : [];
  if (sourceViews.length) {
    workbook.windows.getItemAt(0).setActiveWorksheet(sourceViews[0].activeWorksheetId);
    if (sourceViews[0].selectedWorksheetIds?.length) workbook.windows.getItemAt(0).setSelectedWorksheets(sourceViews[0].selectedWorksheetIds);
    for (const sourceView of sourceViews.slice(1)) {
      const window = workbook.windows.add({ activeWorksheet: sourceView.activeWorksheetId });
      if (sourceView.selectedWorksheetIds?.length) window.setSelectedWorksheets(sourceView.selectedWorksheetIds);
    }
  }
  const definedNameSlots = (source.definedNames || []).map((wire, index) => ({
    wire,
    definedName: workbook.definedNames.items[index],
    publicSnapshot: definedNameSnapshot(workbook.definedNames.items[index]),
  }));
  const calculationSlot = source.calculation ? { wire: source.calculation, publicSnapshot: calculationSnapshot(workbook.calculation) } : undefined;
  const snapshots = sourceViews.length ? workbookViewSnapshots(workbook) : [];
  const viewSlots = sourceViews.map((wire, index) => ({ wire, publicSnapshot: snapshots[index] }));
  Object.defineProperty(workbook, WORKBOOK_STATE, {
    configurable: true,
    value: {
      source: envelope.source,
      opaqueOpc: envelope.opaqueOpc,
      diagnostics: envelope.diagnostics,
      themeWire: source.theme,
      publicTheme: normalizeXlsxThemeConfig(workbook.theme),
      connectionSlots,
      definedNameSlots,
      calculationSlot,
      viewSlots,
      worksheetSlots,
      dynamicArraySlotsBySheet,
      sourceBoundFormulaSlotsBySheet,
      tablesBySheet,
      imagesBySheet,
      chartsBySheet,
      sparklinesBySheet,
      pivotsBySheet,
      dataTablesBySheet,
    },
    writable: true,
  });
  return workbook;
}

export async function importXlsxWithOfficeKit(input, options = {}) {
  assertCodecOptions(options, new Set(["limits"]), "importXlsxWithOfficeKit");
  const response = await invokeOfficeKit({
    protocolVersion: OFFICE_KIT_PROTOCOL_VERSION,
    operation: CodecOperation.IMPORT_XLSX,
    family: ArtifactFamily.WORKBOOK,
    file: await inputBytes(input),
    limits: codecLimits(options.limits),
  });
  return workbookFromEnvelope(response.artifact);
}
