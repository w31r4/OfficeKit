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

const MAX_METADATA_CHARS = 32_767;

function equalJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function boundedIndex(value, label) {
  const index = Number(value);
  if (!Number.isSafeInteger(index) || index < 0) throw new TypeError(`${label} must be a non-negative safe integer.`);
  return index;
}

function decodeXmlAttribute(value, label) {
  const text = String(value);
  if (/[<"]/u.test(text) || /&(?!amp;|lt;|gt;|quot;|apos;)/u.test(text)) {
    throw new Error(`${label} is not canonically XML escaped.`);
  }
  return text.replace(/&(amp|lt|gt|quot|apos);/g, (_match, entity) => ({
    amp: "&", lt: "<", gt: ">", quot: "\"", apos: "'",
  })[entity]);
}

function normalizeMetadataText(value, label) {
  if (typeof value !== "string" || !value.length || value.length > MAX_METADATA_CHARS || !isXmlSafeText(value)) {
    throw new TypeError(`${label} must contain 1 through ${MAX_METADATA_CHARS} XML-safe characters.`);
  }
  return value;
}

function isXmlSafeText(value) {
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code === 0x09 || code === 0x0a || code === 0x0d) continue;
    if (code < 0x20 || code === 0x7f || code === 0xfffe || code === 0xffff) return false;
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return false;
      index += 1;
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      return false;
    }
  }
  return true;
}

function normalizeMetadata(value, label) {
  const source = typeof value === "string" ? (() => {
    try { return JSON.parse(value); } catch { throw new TypeError(`${label} must be a JSON object.`); }
  })() : value;
  if (!source || typeof source !== "object" || Array.isArray(source)) {
    throw new TypeError(`${label} must be an object with optional title and description.`);
  }
  const unsupported = Object.keys(source).filter((key) => key !== "title" && key !== "description");
  if (unsupported.length) throw new TypeError(`${label} does not support ${unsupported.join(", ")}.`);
  const normalized = {};
  for (const field of ["title", "description"]) {
    if (source[field] != null) normalized[field] = normalizeMetadataText(source[field], `${label}.${field}`);
  }
  return normalized;
}

function rawWordTables(xml, label) {
  const tables = [];
  let open;
  for (const match of String(xml).matchAll(/<w:tbl\b[^>]*>|<\/w:tbl>/g)) {
    if (match[0] === "</w:tbl>") {
      if (!open) throw new Error(`${label} has an unmatched </w:tbl>.`);
      tables.push({ offset: open.offset, xml: String(xml).slice(open.offset, (match.index ?? 0) + match[0].length) });
      open = undefined;
    } else {
      if (open) throw new Error(`${label} has a nested w:tbl; this workflow refuses nested table graphs.`);
      open = { offset: match.index ?? 0 };
    }
  }
  if (open) throw new Error(`${label} has an unclosed w:tbl.`);
  return tables;
}

function directTablePropertyLeaves(inner, label) {
  const leaves = [];
  const stack = [];
  for (const match of String(inner).matchAll(/<\/?[\w:.-]+\b[^>]*>/g)) {
    const markup = match[0];
    const closing = /^<\/([\w:.-]+)\s*>$/.exec(markup);
    if (closing) {
      const current = stack.pop();
      if (!current || current !== closing[1]) throw new Error(`${label} has unbalanced property markup.`);
      continue;
    }
    const opening = /^<([\w:.-]+)\b[^>]*>$/.exec(markup);
    if (!opening) throw new Error(`${label} contains unsupported table-property markup.`);
    const [, name] = opening;
    const selfClosing = /\/>$/.test(markup);
    if (name === "w:tblCaption" || name === "w:tblDescription") {
      if (stack.length || !selfClosing) throw new Error(`${label} has nested or non-leaf ${name} markup.`);
      leaves.push({ name, markup, offset: match.index ?? 0 });
    }
    if (!selfClosing) stack.push(name);
  }
  if (stack.length) throw new Error(`${label} has unclosed table-property markup.`);
  return leaves;
}

function canonicalMetadataLeaf(leaf, label) {
  const match = /^<w:(tblCaption|tblDescription)\s+w:val="([^"]*)"\s*\/>$/.exec(leaf.markup);
  if (!match) throw new Error(`${label} must be a canonical self-closing leaf with exactly one w:val attribute.`);
  const field = match[1] === "tblCaption" ? "title" : "description";
  return { field, value: normalizeMetadataText(decodeXmlAttribute(match[2], `${label} w:val`), `${label} w:val`) };
}

function tableAccessibilityProfile(tableXml, label) {
  const properties = [...String(tableXml).matchAll(/<w:tblPr\b[^>]*\/>|<w:tblPr\b[^>]*>[\s\S]*?<\/w:tblPr>/g)];
  if (properties.length > 1) throw new Error(`${label} has multiple w:tblPr containers.`);
  if (!properties.length) return { metadata: {}, masked: String(tableXml) };
  const markup = properties[0][0];
  const offset = properties[0].index ?? 0;
  const opening = /^<w:tblPr\b([^>]*)>/.exec(markup);
  if (!opening || opening[1].trim()) throw new Error(`${label} has noncanonical w:tblPr markup.`);
  const inner = /^<w:tblPr\b[^>]*>([\s\S]*)<\/w:tblPr>$/.exec(markup)?.[1];
  if (inner === undefined) throw new Error(`${label} has a self-closing w:tblPr container; it cannot carry table metadata.`);
  const leaves = directTablePropertyLeaves(inner, label);
  const metadata = {};
  for (const leaf of leaves) {
    const { field, value } = canonicalMetadataLeaf(leaf, `${label} ${leaf.name}`);
    if (Object.hasOwn(metadata, field)) throw new Error(`${label} has duplicate ${leaf.name} metadata.`);
    metadata[field] = value;
  }
  let cursor = 0;
  let maskedInner = "";
  for (const leaf of leaves) {
    maskedInner += `${inner.slice(cursor, leaf.offset)}`;
    cursor = leaf.offset + leaf.markup.length;
  }
  maskedInner += inner.slice(cursor);
  const maskedProperties = `<w:tblPr>${maskedInner}</w:tblPr>`;
  return { metadata, masked: `${String(tableXml).slice(0, offset)}${maskedProperties}${String(tableXml).slice(offset + markup.length)}` };
}

function normalizedTargetTableXml(xml, tableOrdinal, expectedTableCount, expectedMetadata, label) {
  const tables = rawWordTables(xml, label);
  if (tables.length !== expectedTableCount) {
    throw new Error(`${label} has ${tables.length} flat native tables, but import exposed ${expectedTableCount}; refusing an ambiguous source-to-model mapping.`);
  }
  const target = tables[tableOrdinal];
  if (!target) throw new Error(`${label} has no native table at ordinal ${tableOrdinal}.`);
  const profile = tableAccessibilityProfile(target.xml, label);
  if (!equalJson(profile.metadata, expectedMetadata)) {
    throw new Error(`${label} native table accessibility metadata does not match the bound source value.`);
  }
  return {
    metadata: profile.metadata,
    normalized: canonicalizeXmlForResidual(
      `${String(xml).slice(0, target.offset)}${profile.masked}${String(xml).slice(target.offset + target.xml.length)}`,
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
    values: block.values.map((row) => [...row]),
    cells: structuredClone(block.cells || []),
    headerRowCount: block.headerRowCount,
    keepTogetherRows: [...block.keepTogetherRows],
    minimumRowHeightsDxa: [...block.minimumRowHeightsDxa],
    widthDxa: block.widthDxa,
    indentDxa: block.indentDxa,
    columnWidthsDxa: [...block.columnWidthsDxa],
    cellMarginsDxa: { ...block.cellMarginsDxa },
    borderColor: block.borderColor,
    borderSize: block.borderSize,
    headerFill: block.headerFill,
    horizontalAlignment: block.horizontalAlignment,
    verticalAlignment: block.verticalAlignment,
    accessibility: normalizeMetadata(block.accessibility || {}, `table ${block.id} accessibility`),
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

function selectTable(document, expectedMetadata, tableBlockIndex) {
  const blockIndex = boundedIndex(tableBlockIndex, "tableBlockIndex");
  const block = document.blocks[blockIndex];
  if (!block || block.kind !== "table") throw new Error("tableBlockIndex does not identify an imported table block.");
  if (block.sourceBound !== true || document.resolve(block.id) !== block) {
    throw new Error("Selected table is not a stable imported source-bound table.");
  }
  const actual = normalizeMetadata(block.accessibility || {}, "selected table accessibility");
  if (!equalJson(actual, expectedMetadata)) {
    throw new Error(`Selected table accessibility metadata does not match the expected source value: expected ${JSON.stringify(expectedMetadata)}, observed ${JSON.stringify(actual)}.`);
  }
  const tableOrdinal = document.blocks.slice(0, blockIndex + 1).filter((item) => item.kind === "table").length - 1;
  return { block, blockIndex, tableOrdinal, snapshot: tableSnapshot(block, blockIndex, tableOrdinal) };
}

async function modelRender(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

/**
 * Changes the non-visible title and/or description of one canonical imported
 * DOCX table through the public DocumentFile path. This is table alternative
 * text, not a visible caption paragraph or a visual table-formatting edit.
 */
export async function editImportedTableAccessibility({
  inputPath,
  outputPath,
  auditPath,
  tableBlockIndex,
  expectedAccessibility,
  replacementAccessibility,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const expected = normalizeMetadata(expectedAccessibility, "expectedAccessibility");
  const replacement = normalizeMetadata(replacementAccessibility, "replacementAccessibility");
  if (equalJson(expected, replacement)) throw new Error("replacementAccessibility must differ from expectedAccessibility.");
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const sourceTables = tableProjection(document);
  const selected = selectTable(document, expected, tableBlockIndex);
  const sourceXml = await readPackagePartText(source, "word/document.xml", "Source DOCX package");
  const sourceResidual = normalizedTargetTableXml(sourceXml, selected.tableOrdinal, sourceTables.length, expected, "source target table");
  selected.block.setAccessibilityMetadata({
    title: replacement.title ?? null,
    description: replacement.description ?? null,
  });
  const sourceVerification = document.verify({ visualQa: true });
  if (!sourceVerification.ok) throw new Error(`Replacement table accessibility metadata fails document verification: ${sourceVerification.ndjson}`);

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await DocumentFile.exportDocx(document);
    await fs.writeFile(temporaryPath, Buffer.from(await exported.arrayBuffer()), { flag: "wx" });
    const output = await fs.readFile(temporaryPath);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source DOCX changed during the transaction; refusing publication.");
    const changed = await changedParts(source, output, "Source-bound table accessibility edit");
    if (!equalJson(changed, ["word/document.xml"])) {
      throw new Error(`Source-bound table accessibility edit changed an unexpected package scope: ${changed.join(", ") || "none"}.`);
    }
    const outputXml = await readPackagePartText(output, "word/document.xml", "Output DOCX package");
    const outputResidual = normalizedTargetTableXml(outputXml, selected.tableOrdinal, sourceTables.length, replacement, "output target table");
    if (outputResidual.normalized !== sourceResidual.normalized) {
      throw new Error("Table accessibility edit changed word/document.xml outside the bound tblCaption/tblDescription leaves.");
    }

    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const afterTables = tableProjection(reimported);
    const expectedTables = structuredClone(sourceTables);
    const expectedTable = expectedTables.find((table) => table.id === selected.snapshot.id);
    if (!expectedTable) throw new Error("Selected table disappeared from the imported table projection.");
    expectedTable.accessibility = replacement;
    if (!equalJson(afterTables, expectedTables)) {
      throw new Error("DOCX export changed imported table identity or semantics outside the requested alternative-text metadata.");
    }
    const roundTrip = selectTable(reimported, replacement, selected.blockIndex);
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
        type: "source-bound-table-accessibility-edit",
        target: { id: selected.snapshot.id, blockIndex: selected.blockIndex, tableOrdinal: selected.tableOrdinal },
        expectedAccessibility: expected,
        replacementAccessibility: replacement,
        retained: {
          name: selected.snapshot.name,
          styleId: selected.snapshot.styleId,
          rows: selected.snapshot.rows,
          columns: selected.snapshot.columns,
          gridColumns: selected.snapshot.gridColumns,
          values: selected.snapshot.values,
        },
      },
      validation: {
        changedParts: changed,
        tableAccessibilityXmlResidual: {
          ok: true,
          tableOrdinal: selected.tableOrdinal,
          normalizedSha256: sha256(Buffer.from(sourceResidual.normalized, "utf8")),
        },
        reimport: { ok: true, tableId: roundTrip.snapshot.id, tableBlockIndex: roundTrip.blockIndex, tableOrdinal: roundTrip.tableOrdinal, accessibility: replacement },
        verify: { ok: true },
        modelRender: { ok: true, ...render },
        nativeRenderRequired: true,
      },
      warnings: ["This transaction changes non-visible Word table alternative-text metadata only. It does not create a visible caption, infer author intent, or change layout; inspect a native render and review the title and description before delivery."],
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

export function parseTableAccessibilityEditCli(argv) {
  const [inputPath, outputPath, auditPath, tableBlockIndex, expectedAccessibility, replacementAccessibility] = argv;
  return {
    inputPath,
    outputPath,
    auditPath,
    tableBlockIndex: boundedIndex(tableBlockIndex, "tableBlockIndex"),
    expectedAccessibility: normalizeMetadata(expectedAccessibility, "expectedAccessibility"),
    replacementAccessibility: normalizeMetadata(replacementAccessibility, "replacementAccessibility"),
  };
}

export function tableAccessibilityCliOutput(result) {
  return {
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    changedParts: result.audit.validation.changedParts,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editImportedTableAccessibility(parseTableAccessibilityEditCli(process.argv.slice(2)));
  console.log(JSON.stringify(tableAccessibilityCliOutput(result)));
}
