import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { DocumentFile, FileBlob } from "office-kit";
import {
  DOCX_MIME,
  assertAbsent,
  canonicalizeXmlForResidual,
  changedParts,
  packageVersion,
  publishNoReplace,
  readPackagePartText,
  requiredText,
  sha256,
} from "../artifact_tool/_source_bound_docx.mjs";

const MAX_COLUMNS = 4_096;
const MAX_DXA = 1_000_000;
const DXA_ATTRIBUTES = new Set(["w", "type"]);
const BORDER_ATTRIBUTES = new Set(["val", "color", "sz", "space"]);
const SHADING_ATTRIBUTES = new Set(["val", "color", "fill"]);
const VERTICAL_ALIGNMENT_ATTRIBUTES = new Set(["val"]);
const VERTICAL_ALIGNMENTS = new Set(["top", "center", "bottom"]);
const GRID_COLUMN_ATTRIBUTES = new Set(["w"]);
const CELL_WIDTH_ATTRIBUTES = new Set(["w", "type"]);
const BORDER_NAMES = ["top", "left", "bottom", "right", "insideH", "insideV"];

function equalJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function boundedIndex(value, label) {
  const index = Number(value);
  if (!Number.isSafeInteger(index) || index < 0) throw new TypeError(`${label} must be a non-negative safe integer.`);
  return index;
}

function normalizedInteger(value, label, { positive = false, maximum = MAX_DXA } = {}) {
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed < (positive ? 1 : 0) || parsed > maximum) {
    throw new TypeError(`${label} must be an integer from ${positive ? 1 : 0} through ${maximum}.`);
  }
  return parsed;
}

function normalizedRgb(value, label) {
  const color = String(value || "");
  if (!/^[0-9A-F]{6}$/.test(color)) throw new TypeError(`${label} must be a six-digit uppercase RGB value.`);
  return color;
}

function normalizedVerticalAlignment(value, label) {
  const alignment = String(value || "");
  if (!VERTICAL_ALIGNMENTS.has(alignment)) {
    throw new TypeError(`${label} must be top, center, or bottom.`);
  }
  return alignment;
}

function normalizeMargins(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new TypeError(`${label} must be an object with top, bottom, start, and end margins.`);
  }
  const expected = ["top", "bottom", "start", "end"];
  const actual = Object.keys(value).sort();
  if (!equalJson(actual, [...expected].sort())) {
    throw new TypeError(`${label} must contain exactly top, bottom, start, and end margins.`);
  }
  return {
    top: normalizedInteger(value.top, `${label}.top`),
    bottom: normalizedInteger(value.bottom, `${label}.bottom`),
    start: normalizedInteger(value.start, `${label}.start`),
    end: normalizedInteger(value.end, `${label}.end`),
  };
}

function normalizeDirectFormatting(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new TypeError(`${label} must be a complete direct table-formatting object.`);
  }
  const required = ["indentDxa", "cellMarginsDxa", "borderColor", "borderSize", "headerFill"];
  const allowed = new Set([...required, "verticalAlignment"]);
  const actual = Object.keys(value).sort();
  if (actual.length < required.length || required.some((key) => !Object.hasOwn(value, key)) || actual.some((key) => !allowed.has(key))) {
    throw new TypeError(`${label} must contain indentDxa, cellMarginsDxa, borderColor, borderSize, and headerFill, plus optional verticalAlignment.`);
  }
  const borderSize = normalizedInteger(value.borderSize, `${label}.borderSize`, { maximum: 96 });
  if (borderSize === 1) throw new TypeError(`${label}.borderSize must be zero or from 2 through 96.`);
  return {
    indentDxa: normalizedInteger(value.indentDxa, `${label}.indentDxa`),
    cellMarginsDxa: normalizeMargins(value.cellMarginsDxa, `${label}.cellMarginsDxa`),
    borderColor: normalizedRgb(value.borderColor, `${label}.borderColor`),
    borderSize,
    headerFill: normalizedRgb(value.headerFill, `${label}.headerFill`),
    ...(Object.hasOwn(value, "verticalAlignment") ? { verticalAlignment: normalizedVerticalAlignment(value.verticalAlignment, `${label}.verticalAlignment`) } : {}),
  };
}

function canonicalUnsignedInteger(value, label, { positive = false, maximum = MAX_DXA } = {}) {
  if (!/^(?:0|[1-9]\d*)$/.test(String(value))) throw new Error(`${label} must be a canonical unsigned integer.`);
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed > maximum || (positive && parsed < 1)) {
    throw new Error(`${label} must be an integer from ${positive ? 1 : 0} through ${maximum}.`);
  }
  return parsed;
}

function onlyAttributes(attributes, allowed, label) {
  const unknown = Object.keys(attributes).filter((key) => !allowed.has(key));
  if (unknown.length) throw new Error(`${label} has unsupported attributes: ${unknown.join(", ")}.`);
}

// The source-bound parser deliberately recognizes only canonical self-closing
// WordprocessingML leaves. It is not a generic XML rewriting layer: aliases,
// duplicate attributes, children, and unfamiliar metadata are rejected before
// any owned values are masked.
function wordAttributes(markup, label) {
  const openingMatch = /^<w:[\w.-]+\b([\s\S]*?)\/>$/.exec(String(markup));
  if (!openingMatch) throw new Error(`${label} is not a canonical self-closing WordprocessingML leaf.`);
  const result = {};
  let remaining = openingMatch[1].trim();
  while (remaining) {
    const match = /^([:\w.-]+)="([^"]*)"\s*/.exec(remaining);
    if (!match) throw new Error(`${label} has unsupported XML attribute syntax.`);
    const [, qualifiedName, attributeValue] = match;
    const separator = qualifiedName.indexOf(":");
    if (separator <= 0 || qualifiedName.slice(0, separator) !== "w") {
      throw new Error(`${label} has a noncanonical attribute namespace: ${qualifiedName}.`);
    }
    const localName = qualifiedName.slice(separator + 1);
    if (!localName || Object.hasOwn(result, localName)) {
      throw new Error(`${label} has a duplicate or invalid w: attribute: ${qualifiedName}.`);
    }
    result[localName] = attributeValue;
    remaining = remaining.slice(match[0].length);
  }
  return result;
}

function wordElementInner(markup, name, label) {
  const match = new RegExp(`^<w:${name}\\b([^>]*)>([\\s\\S]*)<\\/w:${name}>$`).exec(String(markup));
  if (!match || match[1].trim()) throw new Error(`${label} is not a canonical <w:${name}> container.`);
  return match[2];
}

function exactlyOne(markups, label) {
  if (markups.length !== 1) throw new Error(`${label} requires exactly one matching leaf; found ${markups.length}.`);
  return markups[0];
}

function selfClosingLeaves(markup, name) {
  return [...String(markup).matchAll(new RegExp(`<w:${name}\\b[^>]*\\/>`, "g"))].map((match) => match[0]);
}

function canonicalDxaLeaf(markup, label, { positive = false } = {}) {
  const attributes = wordAttributes(markup, label);
  onlyAttributes(attributes, DXA_ATTRIBUTES, label);
  if (!Object.hasOwn(attributes, "w") || !Object.hasOwn(attributes, "type")) {
    throw new Error(`${label} requires w:w and w:type.`);
  }
  if (attributes.type !== "dxa") throw new Error(`${label} w:type must be dxa.`);
  return canonicalUnsignedInteger(attributes.w, `${label} w:w`, { positive });
}

function canonicalGridColumn(markup, label) {
  const attributes = wordAttributes(markup, label);
  onlyAttributes(attributes, GRID_COLUMN_ATTRIBUTES, label);
  if (!Object.hasOwn(attributes, "w")) throw new Error(`${label} is missing w:w.`);
  return canonicalUnsignedInteger(attributes.w, `${label} w:w`, { positive: true });
}

function canonicalCellWidth(markup, label) {
  const attributes = wordAttributes(markup, label);
  onlyAttributes(attributes, CELL_WIDTH_ATTRIBUTES, label);
  if (!Object.hasOwn(attributes, "w") || !Object.hasOwn(attributes, "type")) {
    throw new Error(`${label} requires w:w and w:type.`);
  }
  if (attributes.type !== "dxa") throw new Error(`${label} w:type must be dxa.`);
  return canonicalUnsignedInteger(attributes.w, `${label} w:w`, { positive: true });
}

function canonicalBorder(markup, label) {
  const attributes = wordAttributes(markup, label);
  onlyAttributes(attributes, BORDER_ATTRIBUTES, label);
  for (const name of BORDER_ATTRIBUTES) if (!Object.hasOwn(attributes, name)) throw new Error(`${label} is missing w:${name}.`);
  const size = canonicalUnsignedInteger(attributes.sz, `${label} w:sz`, { maximum: 96 });
  if (size === 1 || (size === 0 ? attributes.val !== "nil" : attributes.val !== "single")) {
    throw new Error(`${label} must use w:val="nil" at size 0 or w:val="single" at size 2 through 96.`);
  }
  if (attributes.space !== "0") throw new Error(`${label} w:space must be 0.`);
  return { color: normalizedRgb(attributes.color, `${label} w:color`), size };
}

function canonicalShading(markup, label) {
  const attributes = wordAttributes(markup, label);
  onlyAttributes(attributes, SHADING_ATTRIBUTES, label);
  for (const name of SHADING_ATTRIBUTES) if (!Object.hasOwn(attributes, name)) throw new Error(`${label} is missing w:${name}.`);
  if (attributes.val !== "clear" || attributes.color !== "auto") {
    throw new Error(`${label} must use w:val="clear" and w:color="auto".`);
  }
  return normalizedRgb(attributes.fill, `${label} w:fill`);
}

function canonicalVerticalAlignment(markup, label) {
  const attributes = wordAttributes(markup, label);
  onlyAttributes(attributes, VERTICAL_ALIGNMENT_ATTRIBUTES, label);
  if (!Object.hasOwn(attributes, "val")) throw new Error(`${label} is missing w:val.`);
  return normalizedVerticalAlignment(attributes.val, `${label} w:val`);
}

function rawWordTables(xml, label) {
  const source = String(xml);
  const tables = [];
  let open;
  for (const match of source.matchAll(/<w:tbl\b[^>]*>|<\/w:tbl>/g)) {
    if (match[0] === "</w:tbl>") {
      if (!open) throw new Error(`${label} has an unmatched </w:tbl>.`);
      tables.push({ index: tables.length, offset: open.offset, xml: source.slice(open.offset, (match.index ?? 0) + match[0].length) });
      open = undefined;
    } else {
      if (open) throw new Error(`${label} has a nested w:tbl; this flat-table workflow refuses nested table graphs.`);
      open = { offset: match.index ?? 0 };
    }
  }
  if (open) throw new Error(`${label} has an unclosed w:tbl.`);
  return tables;
}

function tableWidthsMarkup(tableXml, label) {
  const grids = [...String(tableXml).matchAll(/<w:tblGrid\b[^>]*>[\s\S]*?<\/w:tblGrid>/g)];
  const gridMarkup = exactlyOne(grids.map((match) => match[0]), `${label} w:tblGrid`);
  const gridInner = wordElementInner(gridMarkup, "tblGrid", `${label} w:tblGrid`);
  const gridLeaves = selfClosingLeaves(gridInner, "gridCol");
  if (!gridLeaves.length || gridLeaves.length > MAX_COLUMNS || gridInner.replace(/<w:gridCol\b[^>]*\/>/g, "").trim()) {
    throw new Error(`${label} must contain only 1 through ${MAX_COLUMNS} canonical w:gridCol leaves.`);
  }
  const columnWidthsDxa = gridLeaves.map((markup, index) => canonicalGridColumn(markup, `${label} grid column ${index}`));
  const cellLeaves = selfClosingLeaves(tableXml, "tcW");
  if (!cellLeaves.length) throw new Error(`${label} has no canonical w:tcW leaves.`);
  return { columnWidthsDxa, cellWidthsDxa: cellLeaves.map((markup, index) => canonicalCellWidth(markup, `${label} cell width ${index}`)) };
}

function expectedCellWidths(columnWidthsDxa, rows) {
  return Array.from({ length: rows }, () => columnWidthsDxa).flat();
}

function tablePropertiesMarkup(tableXml, label) {
  const properties = [...String(tableXml).matchAll(/<w:tblPr\b[^>]*>[\s\S]*?<\/w:tblPr>/g)].map((match) => match[0]);
  const markup = exactlyOne(properties, `${label} w:tblPr`);
  const inner = wordElementInner(markup, "tblPr", `${label} w:tblPr`);
  const width = exactlyOne(selfClosingLeaves(inner, "tblW"), `${label} w:tblW`);
  const indent = exactlyOne(selfClosingLeaves(inner, "tblInd"), `${label} w:tblInd`);
  const layout = exactlyOne(selfClosingLeaves(inner, "tblLayout"), `${label} w:tblLayout`);
  const borders = exactlyOne([...inner.matchAll(/<w:tblBorders\b[^>]*>[\s\S]*?<\/w:tblBorders>/g)].map((match) => match[0]), `${label} w:tblBorders`);
  const margins = exactlyOne([...inner.matchAll(/<w:tblCellMar\b[^>]*>[\s\S]*?<\/w:tblCellMar>/g)].map((match) => match[0]), `${label} w:tblCellMar`);
  const stripped = inner
    .replace(/<w:tblStyle\b[^>]*\/>/g, "")
    .replace(width, "")
    .replace(indent, "")
    .replace(layout, "")
    .replace(borders, "")
    .replace(margins, "");
  if (stripped.trim()) throw new Error(`${label} has unsupported direct table-property markup.`);
  const layoutAttributes = wordAttributes(layout, `${label} w:tblLayout`);
  onlyAttributes(layoutAttributes, new Set(["type"]), `${label} w:tblLayout`);
  if (layoutAttributes.type !== "fixed") throw new Error(`${label} w:tblLayout must be fixed.`);
  return {
    markup,
    widthMarkup: width,
    indentMarkup: indent,
    bordersMarkup: borders,
    marginsMarkup: margins,
    widthDxa: canonicalDxaLeaf(width, `${label} w:tblW`, { positive: true }),
    indentDxa: canonicalDxaLeaf(indent, `${label} w:tblInd`),
  };
}

function tableBordersMarkup(markup, label) {
  const inner = wordElementInner(markup, "tblBorders", label);
  const leaves = [...inner.matchAll(/<w:(top|left|bottom|right|insideH|insideV)\b[^>]*\/>/g)];
  if (leaves.length !== BORDER_NAMES.length || inner.replace(/<w:(top|left|bottom|right|insideH|insideV)\b[^>]*\/>/g, "").trim()) {
    throw new Error(`${label} must contain exactly six canonical uniform border leaves.`);
  }
  const values = new Map();
  for (const match of leaves) {
    const name = /^<w:([\w.-]+)\b/.exec(match[0])?.[1];
    if (!name || values.has(name)) throw new Error(`${label} has duplicate or invalid border leaves.`);
    values.set(name, canonicalBorder(match[0], `${label} w:${name}`));
  }
  if (values.size !== BORDER_NAMES.length || BORDER_NAMES.some((name) => !values.has(name))) {
    throw new Error(`${label} is missing a required border leaf.`);
  }
  const first = values.get("top");
  if (!first || [...values.values()].some((value) => value.color !== first.color || value.size !== first.size)) {
    throw new Error(`${label} must use one uniform border color and size.`);
  }
  return { color: first.color, size: first.size };
}

function tableMarginsMarkup(markup, label) {
  const inner = wordElementInner(markup, "tblCellMar", label);
  const leaves = [...inner.matchAll(/<w:(top|bottom|start|end|left|right)\b[^>]*\/>/g)];
  if (leaves.length !== 4 || inner.replace(/<w:(top|bottom|start|end|left|right)\b[^>]*\/>/g, "").trim()) {
    throw new Error(`${label} must contain exactly four canonical cell-margin leaves.`);
  }
  const values = new Map();
  for (const match of leaves) {
    const name = /^<w:([\w.-]+)\b/.exec(match[0])?.[1];
    if (!name || values.has(name)) throw new Error(`${label} has duplicate or invalid cell-margin leaves.`);
    values.set(name, canonicalDxaLeaf(match[0], `${label} w:${name}`));
  }
  if (!values.has("top") || !values.has("bottom") || (values.has("start") === values.has("left")) ||
      (values.has("end") === values.has("right"))) {
    throw new Error(`${label} must contain top, bottom, and exactly one start/end margin spelling.`);
  }
  return {
    top: values.get("top"),
    bottom: values.get("bottom"),
    start: values.get("start") ?? values.get("left"),
    end: values.get("end") ?? values.get("right"),
  };
}

function headerRowMarkup(tableXml, expectedRows, expectedColumns, label) {
  const rows = [...String(tableXml).matchAll(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g)].map((match) => match[0]);
  if (rows.length !== expectedRows) throw new Error(`${label} has ${rows.length} raw rows, but import exposed ${expectedRows}.`);
  const markup = rows[0];
  const cells = [...markup.matchAll(/<w:tc\b[^>]*>[\s\S]*?<\/w:tc>/g)].map((match) => match[0]);
  if (cells.length !== expectedColumns) throw new Error(`${label} has ${cells.length} raw header cells, but import exposed ${expectedColumns}.`);
  const fills = [];
  for (let index = 0; index < cells.length; index += 1) {
    const properties = [...cells[index].matchAll(/<w:tcPr\b[^>]*>[\s\S]*?<\/w:tcPr>/g)].map((match) => match[0]);
    const cellProperties = exactlyOne(properties, `${label} header cell ${index} w:tcPr`);
    const inner = wordElementInner(cellProperties, "tcPr", `${label} header cell ${index} w:tcPr`);
    const shading = exactlyOne(selfClosingLeaves(inner, "shd"), `${label} header cell ${index} w:shd`);
    fills.push(canonicalShading(shading, `${label} header cell ${index} w:shd`));
  }
  if (fills.some((fill) => fill !== fills[0])) throw new Error(`${label} header cells must use one uniform fill.`);
  return { markup, fill: fills[0] };
}

function tableVerticalAlignment(tableXml, expectedRows, expectedColumns, label) {
  const rows = [...String(tableXml).matchAll(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g)].map((match) => match[0]);
  if (rows.length !== expectedRows) throw new Error(`${label} has ${rows.length} raw rows, but import exposed ${expectedRows}.`);
  let initialized = false;
  let alignment;
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const cells = [...rows[rowIndex].matchAll(/<w:tc\b[^>]*>[\s\S]*?<\/w:tc>/g)].map((match) => match[0]);
    if (cells.length !== expectedColumns) {
      throw new Error(`${label} row ${rowIndex} has ${cells.length} raw cells, but import exposed ${expectedColumns}.`);
    }
    for (let cellIndex = 0; cellIndex < cells.length; cellIndex += 1) {
      const cellLabel = `${label} cell ${rowIndex},${cellIndex} w:tcPr`;
      const properties = exactlyOne([...cells[cellIndex].matchAll(/<w:tcPr\b[^>]*>[\s\S]*?<\/w:tcPr>/g)].map((match) => match[0]), cellLabel);
      const inner = wordElementInner(properties, "tcPr", cellLabel);
      const leaves = selfClosingLeaves(inner, "vAlign");
      if (/<w:vAlign\b/.test(inner) && leaves.length !== 1) {
        throw new Error(`${cellLabel} must contain zero or one canonical w:vAlign leaf.`);
      }
      const current = leaves.length ? canonicalVerticalAlignment(leaves[0], `${cellLabel} w:vAlign`) : undefined;
      if (!initialized) {
        alignment = current;
        initialized = true;
      } else if (alignment !== current) {
        throw new Error(`${label} must use one uniform physical-cell w:vAlign value, or omit it from every cell.`);
      }
    }
  }
  return alignment;
}

function maskTableVerticalAlignment(tableXml, expectedRows, expectedColumns, label) {
  tableVerticalAlignment(tableXml, expectedRows, expectedColumns, label);
  let propertiesSeen = 0;
  const masked = String(tableXml).replace(/<w:tcPr\b[^>]*>[\s\S]*?<\/w:tcPr>/g, (properties) => {
    const cellIndex = propertiesSeen;
    propertiesSeen += 1;
    const cellLabel = `${label} cell property ${cellIndex} w:tcPr`;
    const inner = wordElementInner(properties, "tcPr", cellLabel);
    const leaves = selfClosingLeaves(inner, "vAlign");
    if (/<w:vAlign\b/.test(inner) && leaves.length !== 1) {
      throw new Error(`${cellLabel} must contain zero or one canonical w:vAlign leaf.`);
    }
    if (leaves.length) canonicalVerticalAlignment(leaves[0], `${cellLabel} w:vAlign`);
    return `<w:tcPr>${inner.replace(/<w:vAlign\b[^>]*\/>/g, "")}<w:vAlign w:val="officeKitVerticalAlignmentMasked"/></w:tcPr>`;
  });
  if (propertiesSeen !== expectedRows * expectedColumns) {
    throw new Error(`${label} has ${propertiesSeen} cell property containers, but import exposed ${expectedRows * expectedColumns}.`);
  }
  return masked;
}

function maskDxaLeaf(markup, label, marker) {
  canonicalDxaLeaf(markup, label);
  const name = /^<w:([\w.-]+)\b/.exec(markup)?.[1];
  if (!name) throw new Error(`${label} has no WordprocessingML leaf name.`);
  return `<w:${name} w:type="dxa" w:w="${marker}"/>`;
}

function maskBorders(markup, label) {
  tableBordersMarkup(markup, label);
  return String(markup).replace(/<w:(top|left|bottom|right|insideH|insideV)\b[^>]*\/>/g, (leaf) => {
    const name = /^<w:([\w.-]+)\b/.exec(leaf)?.[1];
    canonicalBorder(leaf, `${label} w:${name}`);
    return `<w:${name} w:color="officeKitBorderColorMasked" w:space="officeKitBorderSpaceMasked" w:sz="officeKitBorderSizeMasked" w:val="officeKitBorderValueMasked"/>`;
  });
}

function maskMargins(markup, label) {
  tableMarginsMarkup(markup, label);
  return String(markup).replace(/<w:(top|bottom|start|end|left|right)\b[^>]*\/>/g, (leaf) => {
    const name = /^<w:([\w.-]+)\b/.exec(leaf)?.[1];
    return maskDxaLeaf(leaf, `${label} w:${name}`, "officeKitMarginMasked");
  });
}

function maskHeaderRow(markup, expectedColumns, label) {
  headerRowMarkup(`<w:tbl>${markup}</w:tbl>`, 1, expectedColumns, label);
  let propertiesSeen = 0;
  const masked = String(markup).replace(/<w:tcPr\b[^>]*>[\s\S]*?<\/w:tcPr>/g, (properties) => {
    propertiesSeen += 1;
    const inner = wordElementInner(properties, "tcPr", `${label} header cell ${propertiesSeen - 1} w:tcPr`);
    const shading = exactlyOne(selfClosingLeaves(inner, "shd"), `${label} header cell ${propertiesSeen - 1} w:shd`);
    canonicalShading(shading, `${label} header cell ${propertiesSeen - 1} w:shd`);
    return properties.replace(shading, '<w:shd w:color="auto" w:fill="officeKitHeaderFillMasked" w:val="clear"/>');
  });
  if (propertiesSeen !== expectedColumns) throw new Error(`${label} has unexpected header cell property count.`);
  return masked;
}

function tableRawFormatting(tableXml, { rows, columns }, label) {
  const properties = tablePropertiesMarkup(tableXml, label);
  const borders = tableBordersMarkup(properties.bordersMarkup, `${label} w:tblBorders`);
  const margins = tableMarginsMarkup(properties.marginsMarkup, `${label} w:tblCellMar`);
  const header = headerRowMarkup(tableXml, rows, columns, label);
  const verticalAlignment = tableVerticalAlignment(tableXml, rows, columns, label);
  const widths = tableWidthsMarkup(tableXml, label);
  const formatting = {
    indentDxa: properties.indentDxa,
    cellMarginsDxa: margins,
    borderColor: borders.color,
    borderSize: borders.size,
    headerFill: header.fill,
    ...(verticalAlignment === undefined ? {} : { verticalAlignment }),
  };
  const maskedProperties = properties.markup
    .replace(properties.indentMarkup, maskDxaLeaf(properties.indentMarkup, `${label} w:tblInd`, "officeKitTableIndentMasked"))
    .replace(properties.bordersMarkup, maskBorders(properties.bordersMarkup, `${label} w:tblBorders`))
    .replace(properties.marginsMarkup, maskMargins(properties.marginsMarkup, `${label} w:tblCellMar`));
  const withMaskedProperties = String(tableXml).replace(properties.markup, maskedProperties);
  const currentHeader = headerRowMarkup(withMaskedProperties, rows, columns, `${label} masked table`).markup;
  const withMaskedHeader = withMaskedProperties.replace(currentHeader, maskHeaderRow(currentHeader, columns, `${label} header`));
  const masked = maskTableVerticalAlignment(withMaskedHeader, rows, columns, `${label} cell alignment`);
  return {
    widthDxa: properties.widthDxa,
    formatting,
    columnWidthsDxa: widths.columnWidthsDxa,
    cellWidthsDxa: widths.cellWidthsDxa,
    masked,
  };
}

function tableSnapshot(block, blockIndex, tableOrdinal) {
  return {
    id: block.id,
    blockIndex,
    tableOrdinal,
    name: block.name || "",
    sourceBound: block.sourceBound === true,
    styleId: block.styleId || "",
    rows: block.rows,
    columns: block.columns,
    gridColumns: block.gridColumns,
    values: block.values.map((row) => [...row]),
    cells: block.cells?.map((cell) => ({
      row: cell.row,
      column: cell.column,
      gridColumn: cell.gridColumn,
      columnSpan: cell.columnSpan,
      rowSpan: cell.rowSpan,
      verticalMerge: cell.verticalMerge,
      editable: cell.editable,
      textPatchable: cell.textPatchable,
      contentControl: cell.contentControl ? structuredClone(cell.contentControl) : undefined,
    })),
    widthDxa: block.widthDxa,
    indentDxa: block.indentDxa,
    columnWidthsDxa: [...block.columnWidthsDxa],
    cellMarginsDxa: { ...block.cellMarginsDxa },
    borderColor: block.borderColor,
    borderSize: block.borderSize,
    headerFill: block.headerFill,
    ...(block.verticalAlignment == null ? {} : { verticalAlignment: block.verticalAlignment }),
  };
}

function tableProjection(document) {
  let tableOrdinal = 0;
  return document.blocks.flatMap((block, blockIndex) => {
    if (block.kind !== "table") return [];
    const snapshot = tableSnapshot(block, blockIndex, tableOrdinal);
    tableOrdinal += 1;
    return [snapshot];
  });
}

function assertFixedRectangularTable(block, label) {
  if (block.sourceBound !== true) throw new Error(`The ${label} is not an imported source-bound table.`);
  if (!Array.isArray(block.values) || !Number.isSafeInteger(block.rows) || !Number.isSafeInteger(block.columns) ||
      block.rows < 1 || block.columns < 1 || block.rows > MAX_COLUMNS || block.columns > MAX_COLUMNS ||
      block.values.length !== block.rows || block.values.some((row) => !Array.isArray(row) || row.length !== block.columns)) {
    throw new Error(`The ${label} is not a bounded rectangular table.`);
  }
  const widths = block.columnWidthsDxa?.map((value, index) => normalizedInteger(value, `${label}.columnWidthsDxa[${index}]`, { positive: true }));
  if (!Array.isArray(widths) || widths.length !== block.columns || block.gridColumns !== block.columns ||
      widths.reduce((sum, value) => sum + value, 0) !== block.widthDxa) {
    throw new Error(`The ${label} does not expose a complete fixed-layout table-width profile.`);
  }
  if (!Array.isArray(block.cells) || block.cells.length !== block.rows * block.columns || block.textPatches?.length) {
    throw new Error(`The ${label} must have one unpatched physical cell per rectangular grid position.`);
  }
  const cells = new Map(block.cells.map((cell) => [`${cell.row}:${cell.column}`, cell]));
  for (let row = 0; row < block.rows; row += 1) {
    for (let column = 0; column < block.columns; column += 1) {
      const cell = cells.get(`${row}:${column}`);
      if (!cell || cell.gridColumn !== column || cell.columnSpan !== 1 || cell.rowSpan !== 1 ||
          cell.verticalMerge !== "none" || cell.editable !== true || cell.contentControl) {
        throw new Error(`The ${label} has merged, read-only, content-control, or irregular cell geometry at ${row},${column}.`);
      }
    }
  }
  return widths;
}

function blockFormatting(block, label) {
  return normalizeDirectFormatting({
    indentDxa: block.indentDxa,
    cellMarginsDxa: block.cellMarginsDxa,
    borderColor: block.borderColor,
    borderSize: block.borderSize,
    headerFill: block.headerFill,
    ...(block.verticalAlignment == null ? {} : { verticalAlignment: block.verticalAlignment }),
  }, label);
}

function selectCanonicalTable(document, { tableBlockIndex, expectedFormatting }) {
  const blockIndex = boundedIndex(tableBlockIndex, "tableBlockIndex");
  const block = document.blocks[blockIndex];
  if (!block) throw new Error("tableBlockIndex is outside the imported document.");
  if (block.kind !== "table") throw new Error("tableBlockIndex does not identify an imported table block.");
  if (document.resolve(block.id) !== block) throw new Error("Selected table locator did not resolve to the inspected object.");
  assertFixedRectangularTable(block, "selected table");
  const actual = blockFormatting(block, "selected table formatting");
  if (!equalJson(actual, expectedFormatting)) {
    throw new Error(`Selected table direct formatting does not match the expected source value: expected ${JSON.stringify(expectedFormatting)}, observed ${JSON.stringify(actual)}.`);
  }
  const tableOrdinal = document.blocks.slice(0, blockIndex + 1).filter((item) => item.kind === "table").length - 1;
  return { block, blockIndex, tableOrdinal, snapshot: tableSnapshot(block, blockIndex, tableOrdinal) };
}

function normalizedTargetTableFormattingXml(xml, selected, expectedTableCount, label) {
  const tables = rawWordTables(xml, label);
  if (tables.length !== expectedTableCount) {
    throw new Error(`${label} has ${tables.length} flat native tables, but import exposed ${expectedTableCount}; refusing an ambiguous source-to-model mapping.`);
  }
  const table = tables[selected.tableOrdinal];
  if (!table) throw new Error(`${label} has no native table at ordinal ${selected.tableOrdinal}.`);
  const raw = tableRawFormatting(table.xml, { rows: selected.block.rows, columns: selected.block.columns }, label);
  return {
    ...raw,
    tableCount: tables.length,
    normalized: canonicalizeXmlForResidual(
      `${xml.slice(0, table.offset)}${raw.masked}${xml.slice(table.offset + table.xml.length)}`,
      label,
    ),
  };
}

function assertRawMatchesModel(raw, selected, expectedFormatting, label) {
  const sourceWidths = selected.snapshot.columnWidthsDxa;
  if (raw.widthDxa !== selected.snapshot.widthDxa || !equalJson(raw.columnWidthsDxa, sourceWidths) ||
      !equalJson(raw.cellWidthsDxa, expectedCellWidths(sourceWidths, selected.snapshot.rows))) {
    throw new Error(`${label} table-width leaves do not match the complete inspected fixed-layout profile.`);
  }
  if (!equalJson(raw.formatting, expectedFormatting)) {
    throw new Error(`${label} direct formatting leaves do not match the expected source profile.`);
  }
}

function applyFormatting(block, formatting) {
  block.indentDxa = formatting.indentDxa;
  block.cellMarginsDxa = { ...formatting.cellMarginsDxa };
  block.borderColor = formatting.borderColor;
  block.borderSize = formatting.borderSize;
  block.headerFill = formatting.headerFill;
  if (formatting.verticalAlignment === undefined) delete block.verticalAlignment;
  else block.verticalAlignment = formatting.verticalAlignment;
}

function assertValidReplacement(document) {
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) throw new Error(`Replacement table formatting fails document verification: ${verification.ndjson}`);
}

async function modelRender(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

/**
 * Changes one complete direct formatting profile on an imported fixed-layout
 * DOCX table through the public DocumentFile path. Table width/grid/cell-width
 * leaves, table text/style/topology, and every other package part are bound;
 * only the canonical indent, six uniform borders, four cell margins, first-row
 * cell-shading fills, and uniform physical-cell w:vAlign leaves in
 * word/document.xml may differ.
 */
export async function editImportedTableFormatting({
  inputPath,
  outputPath,
  auditPath,
  tableBlockIndex,
  expectedFormatting,
  replacementFormatting,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const expected = normalizeDirectFormatting(expectedFormatting, "expectedFormatting");
  const replacement = normalizeDirectFormatting(replacementFormatting, "replacementFormatting");
  if (equalJson(expected, replacement)) throw new Error("replacementFormatting must differ from expectedFormatting.");
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const sourceTables = tableProjection(document);
  const selected = selectCanonicalTable(document, { tableBlockIndex, expectedFormatting: expected });
  const sourceXml = await readPackagePartText(source, "word/document.xml", "Source DOCX package");
  const sourceResidual = normalizedTargetTableFormattingXml(sourceXml, selected, sourceTables.length, "source target table");
  assertRawMatchesModel(sourceResidual, selected, expected, "Source raw");
  applyFormatting(selected.block, replacement);
  assertValidReplacement(document);

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await DocumentFile.exportDocx(document);
    await fs.writeFile(temporaryPath, Buffer.from(await exported.arrayBuffer()), { flag: "wx" });
    const output = await fs.readFile(temporaryPath);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source DOCX changed during the transaction; refusing publication.");
    const changed = await changedParts(source, output, "Source-bound table formatting edit");
    if (!equalJson(changed, ["word/document.xml"])) {
      throw new Error(`Source-bound table formatting edit changed an unexpected package scope: ${changed.join(", ") || "none"}.`);
    }
    const outputXml = await readPackagePartText(output, "word/document.xml", "Output DOCX package");
    const outputResidual = normalizedTargetTableFormattingXml(outputXml, selected, sourceTables.length, "output target table");
    assertRawMatchesModel(outputResidual, selected, replacement, "Output raw");
    if (sourceResidual.tableCount !== outputResidual.tableCount || outputResidual.normalized !== sourceResidual.normalized) {
      throw new Error("Table formatting edit changed word/document.xml outside the bound indent, borders, cell margins, header fills, and physical-cell alignment leaves.");
    }

    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const roundTrip = selectCanonicalTable(reimported, { tableBlockIndex: selected.blockIndex, expectedFormatting: replacement });
    const afterTables = tableProjection(reimported);
    const expectedTables = structuredClone(sourceTables);
    const expectedTable = expectedTables.find((table) => table.id === selected.snapshot.id);
    if (!expectedTable) throw new Error("Selected table disappeared from the imported table projection.");
    expectedTable.indentDxa = replacement.indentDxa;
    expectedTable.cellMarginsDxa = { ...replacement.cellMarginsDxa };
    expectedTable.borderColor = replacement.borderColor;
    expectedTable.borderSize = replacement.borderSize;
    expectedTable.headerFill = replacement.headerFill;
    if (replacement.verticalAlignment === undefined) delete expectedTable.verticalAlignment;
    else expectedTable.verticalAlignment = replacement.verticalAlignment;
    if (!equalJson(afterTables, expectedTables)) {
      throw new Error("DOCX export changed imported table identity or semantics outside the requested direct-formatting profile.");
    }
    if (roundTrip.snapshot.id !== selected.snapshot.id) throw new Error("Second import did not preserve the selected table identity.");
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error(`Document verification failed: ${verification.ndjson}`);
    const render = await modelRender(reimported);
    const audit = {
      schema: "office-kit.docx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sourceHash, bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "office-kit", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite", noReplace: true },
      operation: {
        type: "source-bound-table-formatting-edit",
        target: { id: selected.snapshot.id, blockIndex: selected.blockIndex, tableOrdinal: selected.tableOrdinal },
        sourceFormatting: expected,
        replacementFormatting: replacement,
        retainedWidthDxa: selected.snapshot.widthDxa,
        retainedColumnWidthsDxa: selected.snapshot.columnWidthsDxa,
      },
      validation: {
        changedParts: changed,
        tableFormattingXmlResidual: {
          ok: true,
          tableOrdinal: selected.tableOrdinal,
          normalizedSha256: sha256(Buffer.from(sourceResidual.normalized, "utf8")),
        },
        reimport: {
          ok: true,
          tableId: roundTrip.snapshot.id,
          tableBlockIndex: roundTrip.blockIndex,
          tableOrdinal: roundTrip.tableOrdinal,
          formatting: replacement,
        },
        verify: { ok: true },
        modelRender: { ok: true, ...render },
        nativeRenderRequired: true,
      },
      warnings: ["This transaction changes one complete recognized direct table-formatting profile. It retains table width/grid/cell widths and does not calculate Word wrapping or page flow; inspect a native Word or LibreOffice render before delivery."],
    };
    await fs.writeFile(temporaryAuditPath, `${JSON.stringify(audit, null, 2)}\n`, { flag: "wx" });
    await publishNoReplace(temporaryPath, finalPath);
    try {
      await publishNoReplace(temporaryAuditPath, finalAuditPath);
    } catch (error) {
      await fs.rm(finalPath, { force: true });
      throw error;
    }
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } catch (error) {
    await Promise.all([fs.rm(temporaryPath, { force: true }), fs.rm(temporaryAuditPath, { force: true })]);
    throw error;
  }
}

function parseJsonFormatting(value, label) {
  try {
    return JSON.parse(requiredText(value, label));
  } catch (error) {
    throw new TypeError(`${label} must be JSON for one complete table-formatting object: ${error.message}`);
  }
}

export function parseTableFormattingEditCli(argv) {
  const [inputPath, outputPath, auditPath, tableBlockIndex, expectedFormatting, replacementFormatting] = argv;
  return {
    inputPath,
    outputPath,
    auditPath,
    tableBlockIndex: boundedIndex(tableBlockIndex, "tableBlockIndex"),
    expectedFormatting: parseJsonFormatting(expectedFormatting, "expectedFormatting"),
    replacementFormatting: parseJsonFormatting(replacementFormatting, "replacementFormatting"),
  };
}

export function tableFormattingCliOutput(result) {
  return {
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    changedParts: result.audit.validation.changedParts,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editImportedTableFormatting(parseTableFormattingEditCli(process.argv.slice(2)));
  console.log(JSON.stringify(tableFormattingCliOutput(result)));
}
