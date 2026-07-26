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

const MAX_ROWS = 4_096;
const ROW_PROPERTY_ORDER = new Map([["gridBefore", 0], ["gridAfter", 1], ["tblHeader", 2]]);

function equalJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function boundedIndex(value, label) {
  const index = Number(value);
  if (!Number.isSafeInteger(index) || index < 0) throw new TypeError(`${label} must be a non-negative safe integer.`);
  return index;
}

function normalizeHeaderRowCount(value, label, rows = MAX_ROWS) {
  const count = Number(value);
  if (!Number.isSafeInteger(count) || count < 0 || count > rows) {
    throw new TypeError(`${label} must be an integer from 0 through ${rows}.`);
  }
  return count;
}

function canonicalUnsignedInteger(value, label) {
  if (!/^(?:0|[1-9]\d*)$/.test(String(value))) throw new Error(`${label} must be a canonical unsigned integer.`);
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed > MAX_ROWS) throw new Error(`${label} must be an integer from 0 through ${MAX_ROWS}.`);
  return parsed;
}

function wordAttributes(markup, label) {
  const openingMatch = /^<w:[\w.-]+\b([\s\S]*?)\/>$/.exec(String(markup));
  if (!openingMatch) throw new Error(`${label} is not a canonical self-closing WordprocessingML leaf.`);
  const attributes = {};
  let remaining = openingMatch[1].trim();
  while (remaining) {
    const match = /^([:\w.-]+)="([^"]*)"\s*/.exec(remaining);
    if (!match) throw new Error(`${label} has unsupported XML attribute syntax.`);
    const [, qualifiedName, value] = match;
    const separator = qualifiedName.indexOf(":");
    if (separator <= 0 || qualifiedName.slice(0, separator) !== "w") {
      throw new Error(`${label} has a noncanonical attribute namespace: ${qualifiedName}.`);
    }
    const name = qualifiedName.slice(separator + 1);
    if (!name || Object.hasOwn(attributes, name)) throw new Error(`${label} has a duplicate or invalid w: attribute: ${qualifiedName}.`);
    attributes[name] = value;
    remaining = remaining.slice(match[0].length);
  }
  return attributes;
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
      if (open) throw new Error(`${label} has a nested w:tbl; this workflow refuses nested table graphs.`);
      open = { offset: match.index ?? 0 };
    }
  }
  if (open) throw new Error(`${label} has an unclosed w:tbl.`);
  return tables;
}

function rawTableRows(tableXml, label) {
  const rows = [];
  for (const match of String(tableXml).matchAll(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g)) {
    const xml = match[0];
    const opening = /^<w:tr\b([^>]*)>/.exec(xml);
    if (!opening || opening[1].trim()) throw new Error(`${label} has a noncanonical w:tr opening tag.`);
    rows.push({ offset: match.index ?? 0, xml });
  }
  if (!rows.length || rows.length > MAX_ROWS) throw new Error(`${label} must contain 1 through ${MAX_ROWS} direct physical rows.`);
  return rows;
}

// The transaction owns only a canonical w:tblHeader prefix. It deliberately
// refuses arbitrary row properties so that adding/removing a repeat-header flag
// never turns this workflow into a general trPr normalizer.
function headerRowProfile(rowXml, label) {
  const properties = [...String(rowXml).matchAll(/<w:trPr\b[^>]*\/>|<w:trPr\b[^>]*>[\s\S]*?<\/w:trPr>/g)];
  if (properties.length > 1) throw new Error(`${label} has multiple w:trPr elements.`);
  if (!properties.length) return { header: false, masked: String(rowXml) };
  const markup = properties[0][0];
  const offset = properties[0].index ?? 0;
  const selfClosing = /^<w:trPr\b([^>]*)\/>$/.exec(markup);
  const container = /^<w:trPr\b([^>]*)>([\s\S]*)<\/w:trPr>$/.exec(markup);
  const attributes = selfClosing?.[1] ?? container?.[1];
  if (attributes === undefined || attributes.trim()) throw new Error(`${label} has noncanonical w:trPr attributes.`);
  const inner = selfClosing ? "" : container?.[2] ?? "";
  const leaves = [...inner.matchAll(/<w:([\w.-]+)\b[^>]*\/>/g)].map((match) => ({ name: match[1], markup: match[0] }));
  if (inner.replace(/<w:(?:gridBefore|gridAfter|tblHeader)\b[^>]*\/>/g, "").trim()) {
    throw new Error(`${label} has unsupported row-property children.`);
  }
  const seen = new Set();
  let previousOrder = -1;
  let header = false;
  for (const leaf of leaves) {
    const order = ROW_PROPERTY_ORDER.get(leaf.name);
    if (order === undefined || seen.has(leaf.name) || order < previousOrder) {
      throw new Error(`${label} has duplicate, unknown, or reordered row-property markup.`);
    }
    seen.add(leaf.name);
    previousOrder = order;
    const attributesForLeaf = wordAttributes(leaf.markup, `${label} w:${leaf.name}`);
    if (leaf.name === "tblHeader") {
      if (Object.keys(attributesForLeaf).length) throw new Error(`${label} w:tblHeader must use the canonical no-w:val form.`);
      header = true;
    } else {
      if (Object.keys(attributesForLeaf).length !== 1 || !Object.hasOwn(attributesForLeaf, "val")) {
        throw new Error(`${label} w:${leaf.name} must contain exactly one w:val attribute.`);
      }
      canonicalUnsignedInteger(attributesForLeaf.val, `${label} w:${leaf.name} w:val`);
    }
  }
  const retained = leaves.filter((leaf) => leaf.name !== "tblHeader").map((leaf) => leaf.markup).join("");
  const maskedProperties = retained ? `<w:trPr>${retained}</w:trPr>` : "";
  return {
    header,
    masked: `${String(rowXml).slice(0, offset)}${maskedProperties}${String(rowXml).slice(offset + markup.length)}`,
  };
}

function headerTableProfile(tableXml, label) {
  const rows = rawTableRows(tableXml, label);
  const profiles = rows.map((row, index) => ({ ...row, ...headerRowProfile(row.xml, `${label} row ${index}`) }));
  let seenNonHeader = false;
  let headerRowCount = 0;
  for (const profile of profiles) {
    if (profile.header) {
      if (seenNonHeader) throw new Error(`${label} has a non-prefix w:tblHeader row.`);
      headerRowCount += 1;
    } else {
      seenNonHeader = true;
    }
  }
  let cursor = 0;
  let masked = "";
  for (const profile of profiles) {
    masked += `${String(tableXml).slice(cursor, profile.offset)}${profile.masked}`;
    cursor = profile.offset + profile.xml.length;
  }
  masked += String(tableXml).slice(cursor);
  return { headerRowCount, rowCount: rows.length, masked };
}

function normalizedTargetTableHeadersXml(xml, tableOrdinal, expectedTableCount, label) {
  const tables = rawWordTables(xml, label);
  if (tables.length !== expectedTableCount) {
    throw new Error(`${label} has ${tables.length} flat native tables, but import exposed ${expectedTableCount}; refusing an ambiguous source-to-model mapping.`);
  }
  const selected = tables[tableOrdinal];
  if (!selected) throw new Error(`${label} has no native table at ordinal ${tableOrdinal}.`);
  const profile = headerTableProfile(selected.xml, label);
  return {
    ...profile,
    tableCount: tables.length,
    normalized: canonicalizeXmlForResidual(
      `${String(xml).slice(0, selected.offset)}${profile.masked}${String(xml).slice(selected.offset + selected.xml.length)}`,
      label,
    ),
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
    headerRowCount: block.headerRowCount,
    values: block.values.map((row) => [...row]),
    cells: structuredClone(block.cells || []),
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
      block.rows < 1 || block.columns < 1 || block.rows > MAX_ROWS || block.columns > MAX_ROWS ||
      block.values.length !== block.rows || block.values.some((row) => !Array.isArray(row) || row.length !== block.columns)) {
    throw new Error(`The ${label} is not a bounded rectangular table.`);
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
}

function selectCanonicalTable(document, { tableBlockIndex, expectedHeaderRowCount }) {
  const blockIndex = boundedIndex(tableBlockIndex, "tableBlockIndex");
  const block = document.blocks[blockIndex];
  if (!block) throw new Error("tableBlockIndex is outside the imported document.");
  if (block.kind !== "table") throw new Error("tableBlockIndex does not identify an imported table block.");
  if (document.resolve(block.id) !== block) throw new Error("Selected table locator did not resolve to the inspected object.");
  assertFixedRectangularTable(block, "selected table");
  const expected = normalizeHeaderRowCount(expectedHeaderRowCount, "expectedHeaderRowCount", block.rows);
  if (block.headerRowCount !== expected) {
    throw new Error(`Selected table repeat-header count does not match the expected source value: expected ${expected}, observed ${block.headerRowCount}.`);
  }
  const tableOrdinal = document.blocks.slice(0, blockIndex + 1).filter((item) => item.kind === "table").length - 1;
  return { block, blockIndex, tableOrdinal, snapshot: tableSnapshot(block, blockIndex, tableOrdinal) };
}

function assertValidReplacement(document) {
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) throw new Error(`Replacement table repeat headers fail document verification: ${verification.ndjson}`);
}

async function modelRender(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

/**
 * Changes one imported flat rectangular table's leading native w:tblHeader
 * prefix through DocumentFile. Text, styling (including headerFill), grid,
 * cell/row topology, source bytes, and every package part outside
 * word/document.xml remain bound. This changes page-repeat semantics, so a
 * native Word or LibreOffice render remains a required delivery review.
 */
export async function editImportedTableHeaderRows({
  inputPath,
  outputPath,
  auditPath,
  tableBlockIndex,
  expectedHeaderRowCount,
  replacementHeaderRowCount,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const blockIndex = boundedIndex(tableBlockIndex, "tableBlockIndex");
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const sourceTables = tableProjection(document);
  const selected = selectCanonicalTable(document, { tableBlockIndex: blockIndex, expectedHeaderRowCount });
  const expected = normalizeHeaderRowCount(expectedHeaderRowCount, "expectedHeaderRowCount", selected.block.rows);
  const replacement = normalizeHeaderRowCount(replacementHeaderRowCount, "replacementHeaderRowCount", selected.block.rows);
  if (expected === replacement) throw new Error("replacementHeaderRowCount must differ from expectedHeaderRowCount.");
  const sourceXml = await readPackagePartText(source, "word/document.xml", "Source DOCX package");
  const sourceResidual = normalizedTargetTableHeadersXml(sourceXml, selected.tableOrdinal, sourceTables.length, "source target table");
  if (sourceResidual.rowCount !== selected.block.rows || sourceResidual.headerRowCount !== expected) {
    throw new Error("The raw source table repeat-header profile does not match the inspected source-bound table.");
  }
  selected.block.setHeaderRowCount(replacement);
  assertValidReplacement(document);

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await DocumentFile.exportDocx(document);
    await fs.writeFile(temporaryPath, Buffer.from(await exported.arrayBuffer()), { flag: "wx" });
    const output = await fs.readFile(temporaryPath);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source DOCX changed during the transaction; refusing publication.");
    const changed = await changedParts(source, output, "Source-bound table repeat-header edit");
    if (!equalJson(changed, ["word/document.xml"])) {
      throw new Error(`Source-bound table repeat-header edit changed an unexpected package scope: ${changed.join(", ") || "none"}.`);
    }
    const outputXml = await readPackagePartText(output, "word/document.xml", "Output DOCX package");
    const outputResidual = normalizedTargetTableHeadersXml(outputXml, selected.tableOrdinal, sourceTables.length, "output target table");
    if (outputResidual.rowCount !== selected.block.rows || outputResidual.headerRowCount !== replacement) {
      throw new Error("Exported table repeat-header markers do not match the requested replacement.");
    }
    if (sourceResidual.tableCount !== outputResidual.tableCount || outputResidual.normalized !== sourceResidual.normalized) {
      throw new Error("Table repeat-header edit changed word/document.xml outside the requested canonical w:tblHeader leaves.");
    }

    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const roundTrip = selectCanonicalTable(reimported, { tableBlockIndex: selected.blockIndex, expectedHeaderRowCount: replacement });
    const expectedTables = structuredClone(sourceTables);
    const expectedTable = expectedTables.find((table) => table.id === selected.snapshot.id);
    if (!expectedTable) throw new Error("Selected table disappeared from the imported table projection.");
    expectedTable.headerRowCount = replacement;
    if (!equalJson(tableProjection(reimported), expectedTables)) {
      throw new Error("DOCX export changed imported table identity or semantics outside the requested repeat-header count.");
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
        type: "source-bound-table-header-rows-edit",
        target: { id: selected.snapshot.id, blockIndex: selected.blockIndex, tableOrdinal: selected.tableOrdinal },
        expectedHeaderRowCount: expected,
        replacementHeaderRowCount: replacement,
      },
      validation: {
        changedParts: changed,
        tableHeaderRowsXmlResidual: {
          ok: true,
          tableOrdinal: selected.tableOrdinal,
          normalizedSha256: sha256(Buffer.from(sourceResidual.normalized, "utf8")),
        },
        reimport: { ok: true, tableId: roundTrip.snapshot.id, tableBlockIndex: roundTrip.blockIndex, tableOrdinal: roundTrip.tableOrdinal, headerRowCount: replacement },
        verify: { ok: true },
        modelRender: { ok: true, ...render },
        nativeRenderRequired: true,
      },
      warnings: ["This transaction changes native table-header repetition at page boundaries. It does not change headerFill or calculate pagination; inspect a native Word or LibreOffice render before delivery."],
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

export function parseTableHeaderRowsEditCli(argv) {
  const [inputPath, outputPath, auditPath, tableBlockIndex, expectedHeaderRowCount, replacementHeaderRowCount] = argv;
  return {
    inputPath,
    outputPath,
    auditPath,
    tableBlockIndex: boundedIndex(tableBlockIndex, "tableBlockIndex"),
    expectedHeaderRowCount,
    replacementHeaderRowCount,
  };
}

export function tableHeaderRowsCliOutput(result) {
  return {
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    changedParts: result.audit.validation.changedParts,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editImportedTableHeaderRows(parseTableHeaderRowsEditCli(process.argv.slice(2)));
  console.log(JSON.stringify(tableHeaderRowsCliOutput(result)));
}
