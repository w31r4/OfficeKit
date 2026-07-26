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
const GRID_COLUMN_ATTRIBUTES = new Set(["w"]);
const CELL_WIDTH_ATTRIBUTES = new Set(["w", "type"]);

function equalJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function boundedIndex(value, label) {
  const index = Number(value);
  if (!Number.isSafeInteger(index) || index < 0) throw new TypeError(`${label} must be a non-negative safe integer.`);
  return index;
}

function normalizedPositiveInteger(value, label) {
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed < 1 || parsed > MAX_DXA) {
    throw new TypeError(`${label} must be an integer from 1 through ${MAX_DXA}.`);
  }
  return parsed;
}

function normalizeColumnWidths(value, label) {
  if (!Array.isArray(value) || value.length < 1 || value.length > MAX_COLUMNS) {
    throw new TypeError(`${label} must contain 1 through ${MAX_COLUMNS} widths.`);
  }
  return value.map((width, index) => normalizedPositiveInteger(width, `${label}[${index}]`));
}

function canonicalUnsignedInteger(value, label, { positive = false } = {}) {
  if (!/^(?:0|[1-9]\d*)$/.test(String(value))) throw new Error(`${label} must be a canonical unsigned integer.`);
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed > MAX_DXA || (positive && parsed < 1)) {
    throw new Error(`${label} must be an integer from ${positive ? 1 : 0} through ${MAX_DXA}.`);
  }
  return parsed;
}

function onlyAttributes(attributes, allowed, label) {
  const unknown = Object.keys(attributes).filter((key) => !allowed.has(key));
  if (unknown.length) throw new Error(`${label} has unsupported attributes: ${unknown.join(", ")}.`);
}

// This intentionally accepts only the exact source-bound leaves this workflow
// owns. It is not a general XML attribute parser and rejects namespace aliases,
// duplicate attributes, children, or extension metadata before masking a value.
function wordAttributes(opening = "", label = "WordprocessingML leaf") {
  const tag = String(opening);
  const openingMatch = /^<w:[\w.-]+\b([\s\S]*?)\/>$/.exec(tag);
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

function rawWordTables(xml, label) {
  const source = String(xml);
  const tables = [];
  let open;
  for (const match of source.matchAll(/<w:tbl\b[^>]*>|<\/w:tbl>/g)) {
    if (match[0] === "</w:tbl>") {
      if (!open) throw new Error(`${label} has an unmatched </w:tbl>.`);
      tables.push({
        index: tables.length,
        offset: open.offset,
        xml: source.slice(open.offset, (match.index ?? 0) + match[0].length),
      });
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
  if (grids.length !== 1) throw new Error(`${label} must contain exactly one w:tblGrid; found ${grids.length}.`);
  const gridMarkup = grids[0][0];
  const gridOpening = /^<w:tblGrid\b([^>]*)>/.exec(gridMarkup);
  const gridInner = /^<w:tblGrid\b[^>]*>([\s\S]*)<\/w:tblGrid>$/.exec(gridMarkup)?.[1];
  if (!gridOpening || gridOpening[1].trim() || gridInner === undefined) {
    throw new Error(`${label} has noncanonical w:tblGrid markup.`);
  }
  const gridLeaves = [...gridInner.matchAll(/<w:gridCol\b[^>]*\/>/g)];
  if (!gridLeaves.length || gridLeaves.length > MAX_COLUMNS || gridInner.replace(/<w:gridCol\b[^>]*\/>/g, "").trim()) {
    throw new Error(`${label} must contain only 1 through ${MAX_COLUMNS} canonical w:gridCol leaves.`);
  }
  const columnWidthsDxa = gridLeaves.map((match, index) => canonicalGridColumn(match[0], `${label} grid column ${index}`));
  const cellLeaves = [...String(tableXml).matchAll(/<w:tcW\b[^>]*\/>/g)];
  if (!cellLeaves.length) throw new Error(`${label} has no canonical w:tcW leaves.`);
  const cellWidthsDxa = cellLeaves.map((match, index) => canonicalCellWidth(match[0], `${label} cell width ${index}`));
  return { columnWidthsDxa, cellWidthsDxa };
}

function maskTableWidths(tableXml, label) {
  // Parse first so masking cannot hide a new/unknown property on a target leaf.
  tableWidthsMarkup(tableXml, label);
  return String(tableXml)
    .replace(/<w:gridCol\b[^>]*\/>/g, (markup, index) => {
      canonicalGridColumn(markup, `${label} grid column at byte ${index}`);
      return '<w:gridCol w:w="officeKitWidthMasked"/>';
    })
    .replace(/<w:tcW\b[^>]*\/>/g, (markup, index) => {
      canonicalCellWidth(markup, `${label} cell width at byte ${index}`);
      return '<w:tcW w:type="dxa" w:w="officeKitWidthMasked"/>';
    });
}

function expectedCellWidths(columnWidthsDxa, rows) {
  return Array.from({ length: rows }, () => columnWidthsDxa).flat();
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
  const widths = normalizeColumnWidths(block.columnWidthsDxa, `${label}.columnWidthsDxa`);
  if (widths.length !== block.columns || block.gridColumns !== block.columns ||
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

function selectCanonicalTable(document, { tableBlockIndex, expectedColumnWidthsDxa }) {
  const blockIndex = boundedIndex(tableBlockIndex, "tableBlockIndex");
  const block = document.blocks[blockIndex];
  if (!block) throw new Error("tableBlockIndex is outside the imported document.");
  if (block.kind !== "table") throw new Error("tableBlockIndex does not identify an imported table block.");
  if (document.resolve(block.id) !== block) throw new Error("Selected table locator did not resolve to the inspected object.");
  const actual = assertFixedRectangularTable(block, "selected table");
  if (!equalJson(actual, expectedColumnWidthsDxa)) {
    throw new Error(`Selected table column widths do not match the expected source value: expected ${JSON.stringify(expectedColumnWidthsDxa)}, observed ${JSON.stringify(actual)}.`);
  }
  const tableOrdinal = document.blocks.slice(0, blockIndex + 1).filter((item) => item.kind === "table").length - 1;
  return { block, blockIndex, tableOrdinal, snapshot: tableSnapshot(block, blockIndex, tableOrdinal) };
}

function normalizedTargetTableWidthsXml(xml, tableOrdinal, expectedTableCount, label) {
  const tables = rawWordTables(xml, label);
  if (tables.length !== expectedTableCount) {
    throw new Error(`${label} has ${tables.length} flat native tables, but import exposed ${expectedTableCount}; refusing an ambiguous source-to-model mapping.`);
  }
  const selected = tables[tableOrdinal];
  if (!selected) throw new Error(`${label} has no native table at ordinal ${tableOrdinal}.`);
  const profile = tableWidthsMarkup(selected.xml, label);
  const maskedTable = maskTableWidths(selected.xml, label);
  return {
    ...profile,
    tableCount: tables.length,
    normalized: canonicalizeXmlForResidual(
      `${xml.slice(0, selected.offset)}${maskedTable}${xml.slice(selected.offset + selected.xml.length)}`,
      label,
    ),
  };
}

function assertValidReplacement(document) {
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) throw new Error(`Replacement table widths fail document verification: ${verification.ndjson}`);
}

async function modelRender(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

/**
 * Redistributes the fixed-width columns of one imported flat rectangular DOCX
 * table through the public DocumentFile path. The table's total width, direct
 * formatting profile, text, style, row/cell topology, and every other package
 * part are bound: only canonical w:tblGrid/w:gridCol and matching w:tcW widths
 * in word/document.xml may differ.
 */
export async function editImportedTableColumnWidths({
  inputPath,
  outputPath,
  auditPath,
  tableBlockIndex,
  expectedColumnWidthsDxa,
  replacementColumnWidthsDxa,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const expected = normalizeColumnWidths(expectedColumnWidthsDxa, "expectedColumnWidthsDxa");
  const replacement = normalizeColumnWidths(replacementColumnWidthsDxa, "replacementColumnWidthsDxa");
  if (equalJson(expected, replacement)) throw new Error("replacementColumnWidthsDxa must differ from expectedColumnWidthsDxa.");
  if (replacement.length !== expected.length) throw new Error("replacementColumnWidthsDxa must retain the source column count.");
  const expectedTotal = expected.reduce((sum, value) => sum + value, 0);
  if (replacement.reduce((sum, value) => sum + value, 0) !== expectedTotal) {
    throw new Error("replacementColumnWidthsDxa must retain the source table total width.");
  }
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const sourceTables = tableProjection(document);
  const selected = selectCanonicalTable(document, { tableBlockIndex, expectedColumnWidthsDxa: expected });
  const sourceXml = await readPackagePartText(source, "word/document.xml", "Source DOCX package");
  const sourceResidual = normalizedTargetTableWidthsXml(
    sourceXml,
    selected.tableOrdinal,
    sourceTables.length,
    "source target table",
  );
  if (!equalJson(sourceResidual.columnWidthsDxa, expected) ||
      !equalJson(sourceResidual.cellWidthsDxa, expectedCellWidths(expected, selected.block.rows))) {
    throw new Error("The raw source table grid/cell widths do not match the inspected fixed-layout table profile.");
  }
  selected.block.columnWidthsDxa = [...replacement];
  assertValidReplacement(document);

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await DocumentFile.exportDocx(document);
    await fs.writeFile(temporaryPath, Buffer.from(await exported.arrayBuffer()), { flag: "wx" });
    const output = await fs.readFile(temporaryPath);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source DOCX changed during the transaction; refusing publication.");
    const changed = await changedParts(source, output, "Source-bound table column-width edit");
    if (!equalJson(changed, ["word/document.xml"])) {
      throw new Error(`Source-bound table column-width edit changed an unexpected package scope: ${changed.join(", ") || "none"}.`);
    }

    const outputXml = await readPackagePartText(output, "word/document.xml", "Output DOCX package");
    const outputResidual = normalizedTargetTableWidthsXml(
      outputXml,
      selected.tableOrdinal,
      sourceTables.length,
      "output target table",
    );
    if (!equalJson(outputResidual.columnWidthsDxa, replacement) ||
        !equalJson(outputResidual.cellWidthsDxa, expectedCellWidths(replacement, selected.block.rows))) {
      throw new Error("Exported table grid/cell widths do not match the requested fixed-layout replacement.");
    }
    if (sourceResidual.tableCount !== outputResidual.tableCount || outputResidual.normalized !== sourceResidual.normalized) {
      throw new Error("Table column-width edit changed word/document.xml outside the requested w:tblGrid/w:gridCol and matching w:tcW leaves.");
    }

    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const roundTrip = selectCanonicalTable(reimported, {
      tableBlockIndex: selected.blockIndex,
      expectedColumnWidthsDxa: replacement,
    });
    const afterTables = tableProjection(reimported);
    const expectedTables = structuredClone(sourceTables);
    const expectedTable = expectedTables.find((table) => table.id === selected.snapshot.id);
    if (!expectedTable) throw new Error("Selected table disappeared from the imported table projection.");
    expectedTable.columnWidthsDxa = [...replacement];
    if (!equalJson(afterTables, expectedTables)) {
      throw new Error("DOCX export changed imported table identity or semantics outside the requested column-width profile.");
    }
    if (roundTrip.snapshot.id !== selected.snapshot.id) {
      throw new Error("Second import did not preserve the selected table identity.");
    }
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
        type: "source-bound-table-column-widths-edit",
        target: {
          id: selected.snapshot.id,
          blockIndex: selected.blockIndex,
          tableOrdinal: selected.tableOrdinal,
        },
        tableWidthDxa: expectedTotal,
        sourceColumnWidthsDxa: expected,
        replacementColumnWidthsDxa: replacement,
      },
      validation: {
        changedParts: changed,
        tableWidthXmlResidual: {
          ok: true,
          tableOrdinal: selected.tableOrdinal,
          normalizedSha256: sha256(Buffer.from(sourceResidual.normalized, "utf8")),
        },
        reimport: {
          ok: true,
          tableId: roundTrip.snapshot.id,
          tableBlockIndex: roundTrip.blockIndex,
          tableOrdinal: roundTrip.tableOrdinal,
          columnWidthsDxa: replacement,
        },
        verify: { ok: true },
        modelRender: { ok: true, ...render },
        nativeRenderRequired: true,
      },
      warnings: ["This transaction redistributes a fixed table width only. It does not infer ideal widths or calculate resulting Word text wrapping; inspect a native Word or LibreOffice render before delivery."],
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

function parseJsonWidths(value, label) {
  try {
    return JSON.parse(requiredText(value, label));
  } catch (error) {
    throw new TypeError(`${label} must be JSON for a width array: ${error.message}`);
  }
}

export function parseTableColumnWidthsEditCli(argv) {
  const [inputPath, outputPath, auditPath, tableBlockIndex, expectedColumnWidthsDxa, replacementColumnWidthsDxa] = argv;
  return {
    inputPath,
    outputPath,
    auditPath,
    tableBlockIndex: boundedIndex(tableBlockIndex, "tableBlockIndex"),
    expectedColumnWidthsDxa: parseJsonWidths(expectedColumnWidthsDxa, "expectedColumnWidthsDxa"),
    replacementColumnWidthsDxa: parseJsonWidths(replacementColumnWidthsDxa, "replacementColumnWidthsDxa"),
  };
}

export function tableColumnWidthsCliOutput(result) {
  return {
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    changedParts: result.audit.validation.changedParts,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editImportedTableColumnWidths(parseTableColumnWidthsEditCli(process.argv.slice(2)));
  console.log(JSON.stringify(tableColumnWidthsCliOutput(result)));
}
