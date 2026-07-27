import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import JSZip from "jszip";

import { DocumentFile, DocumentModel, FileBlob } from "office-kit";
import { nativeDocumentRenderStatus, verifyDocumentFile } from "./skill-harness/documents/scripts/workflow.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-table-formatting-workflow-"));
const nativeStatus = nativeDocumentRenderStatus();

function flatTables(xml) {
  return [...xml.matchAll(/<w:tbl\b[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
}

function widths(tableXml, tagName) {
  return [...tableXml.matchAll(new RegExp(`<w:${tagName}\\b[^>]*w:w="(\\d+)"[^>]*\\/>`, "g"))].map((match) => Number(match[1]));
}

function assertFormattingMarkup(tableXml, formatting, rows, columns) {
  assert.match(tableXml, new RegExp(`<w:tblInd\\b(?=[^>]*w:w="${formatting.indentDxa}")(?=[^>]*w:type="dxa")[^>]*/>`));
  for (const [side, value] of Object.entries(formatting.cellMarginsDxa)) {
    assert.match(tableXml, new RegExp(`<w:${side}\\b(?=[^>]*w:w="${value}")(?=[^>]*w:type="dxa")[^>]*/>`));
  }
  for (const name of ["top", "left", "bottom", "right", "insideH", "insideV"]) {
    assert.match(tableXml, new RegExp(`<w:${name}\\b(?=[^>]*w:val="single")(?=[^>]*w:color="${formatting.borderColor}")(?=[^>]*w:sz="${formatting.borderSize}")(?=[^>]*w:space="0")[^>]*/>`));
  }
  const firstRow = /<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/.exec(tableXml)?.[0] || "";
  const fills = [...firstRow.matchAll(/<w:shd\b[^>]*w:val="clear"[^>]*w:color="auto"[^>]*w:fill="([0-9A-F]+)"[^>]*\/>/g)].map((match) => match[1]);
  assert.deepEqual(fills, Array(columns).fill(formatting.headerFill));
  const alignments = [...tableXml.matchAll(/<w:vAlign\b[^>]*w:val="([^"]+)"[^>]*\/>/g)].map((match) => match[1]);
  assert.deepEqual(alignments, formatting.verticalAlignment === undefined ? [] : Array(rows * columns).fill(formatting.verticalAlignment));
}

try {
  const sourceDocument = DocumentModel.create({ name: "Source-bound table formatting edit", blocks: [] });
  sourceDocument.addParagraph("The first table is the source-bound formatting target.");
  sourceDocument.addTable({
    name: "target-table",
    values: [
      ["Quarter", "Revenue", "Margin"],
      ["Q1", "1.2M", "44%"],
      ["Q2", "1.4M", "46%"],
    ],
    widthDxa: 9300,
    indentDxa: 120,
    columnWidthsDxa: [2100, 4500, 2700],
    cellMarginsDxa: { top: 80, bottom: 80, start: 120, end: 120 },
    borderColor: "445566",
    borderSize: 8,
    headerFill: "E2E8F0",
    verticalAlignment: "center",
  });
  sourceDocument.addParagraph("The second table is a raw-XML canary.");
  sourceDocument.addTable({
    name: "sibling-table",
    values: [["Keep", "Unchanged"], ["Scope", "Canary"]],
    widthDxa: 9300,
    indentDxa: 120,
    columnWidthsDxa: [3300, 6000],
    cellMarginsDxa: { top: 80, bottom: 80, start: 120, end: 120 },
    borderColor: "224466",
    borderSize: 6,
    headerFill: "DDEBF7",
  });
  const sourcePath = path.join(outputDir, "source.docx");
  await (await DocumentFile.exportDocx(sourceDocument)).save(sourcePath);
  const sourceBytes = await fs.readFile(sourcePath);
  const sourceImported = await DocumentFile.importDocx(await FileBlob.load(sourcePath));
  const sourceFormatting = {
    indentDxa: 120,
    cellMarginsDxa: { top: 80, bottom: 80, start: 120, end: 120 },
    borderColor: "445566",
    borderSize: 8,
    headerFill: "E2E8F0",
    verticalAlignment: "center",
  };
  const replacementFormatting = {
    indentDxa: 240,
    cellMarginsDxa: { top: 100, bottom: 120, start: 160, end: 180 },
    borderColor: "224466",
    borderSize: 12,
    headerFill: "DDEBF7",
    verticalAlignment: "bottom",
  };
  assert.equal(sourceImported.blocks[1]?.sourceBound, true);
  assert.deepEqual(sourceImported.blocks[1]?.columnWidthsDxa, [2100, 4500, 2700]);
  assert.equal(sourceImported.blocks[1]?.verticalAlignment, "center");

  const workflowPath = path.join(repoRoot, "skills", "documents", "skills", "documents", "examples", "officekit-table-formatting-edit-workflow.mjs");
  const {
    editImportedTableFormatting,
    parseTableFormattingEditCli,
    tableFormattingCliOutput,
  } = await import(workflowPath);
  const outputPath = path.join(outputDir, "output.docx");
  const auditPath = path.join(outputDir, "audit.json");
  const result = await editImportedTableFormatting({
    inputPath: sourcePath,
    outputPath,
    auditPath,
    tableBlockIndex: 1,
    expectedFormatting: sourceFormatting,
    replacementFormatting,
  });
  assert.equal(result.audit.provider.actual, "office-kit");
  assert.equal(result.audit.provider.silentFallback, false);
  assert.deepEqual(result.audit.savePolicy, { strategy: "rewrite", noReplace: true });
  assert.equal(result.audit.operation.type, "source-bound-table-formatting-edit");
  assert.deepEqual(result.audit.operation.target, { id: sourceImported.blocks[1].id, blockIndex: 1, tableOrdinal: 0 });
  assert.deepEqual(result.audit.operation.sourceFormatting, sourceFormatting);
  assert.deepEqual(result.audit.operation.replacementFormatting, replacementFormatting);
  assert.equal(result.audit.operation.retainedWidthDxa, 9300);
  assert.deepEqual(result.audit.operation.retainedColumnWidthsDxa, [2100, 4500, 2700]);
  assert.deepEqual(result.audit.validation.changedParts, ["word/document.xml"]);
  assert.equal(result.audit.validation.tableFormattingXmlResidual.ok, true);
  assert.deepEqual(result.audit.validation.reimport.formatting, replacementFormatting);
  assert.equal(result.audit.validation.nativeRenderRequired, true);
  assert.deepEqual(tableFormattingCliOutput(result).changedParts, ["word/document.xml"]);
  assert.deepEqual(await fs.readFile(sourcePath), sourceBytes);

  const outputBytes = await fs.readFile(outputPath);
  const [sourceZip, outputZip] = await Promise.all([JSZip.loadAsync(sourceBytes), JSZip.loadAsync(outputBytes)]);
  const sourceParts = Object.keys(sourceZip.files).filter((partPath) => !sourceZip.files[partPath].dir).sort();
  assert.deepEqual(Object.keys(outputZip.files).filter((partPath) => !outputZip.files[partPath].dir).sort(), sourceParts);
  for (const partPath of sourceParts) {
    if (partPath === "word/document.xml") continue;
    assert.deepEqual(
      Buffer.from(await outputZip.file(partPath).async("uint8array")),
      Buffer.from(await sourceZip.file(partPath).async("uint8array")),
      `Only word/document.xml may change; ${partPath} drifted.`,
    );
  }
  const [sourceXml, outputXml] = await Promise.all([
    sourceZip.file("word/document.xml").async("text"),
    outputZip.file("word/document.xml").async("text"),
  ]);
  const sourceTables = flatTables(sourceXml);
  const outputTables = flatTables(outputXml);
  assert.equal(sourceTables.length, 2);
  assert.equal(outputTables.length, 2);
  assert.deepEqual(widths(sourceTables[0], "gridCol"), [2100, 4500, 2700]);
  assert.deepEqual(widths(outputTables[0], "gridCol"), [2100, 4500, 2700]);
  assert.deepEqual(widths(outputTables[0], "tcW"), [2100, 4500, 2700, 2100, 4500, 2700, 2100, 4500, 2700]);
  assertFormattingMarkup(sourceTables[0], sourceFormatting, 3, 3);
  assertFormattingMarkup(outputTables[0], replacementFormatting, 3, 3);
  assert.equal(outputTables[1], sourceTables[1]);
  const reimported = await DocumentFile.importDocx(await FileBlob.load(outputPath));
  assert.equal(reimported.blocks[1]?.indentDxa, replacementFormatting.indentDxa);
  assert.deepEqual(reimported.blocks[1]?.cellMarginsDxa, replacementFormatting.cellMarginsDxa);
  assert.equal(reimported.blocks[1]?.borderColor, replacementFormatting.borderColor);
  assert.equal(reimported.blocks[1]?.borderSize, replacementFormatting.borderSize);
  assert.equal(reimported.blocks[1]?.headerFill, replacementFormatting.headerFill);
  assert.equal(reimported.blocks[1]?.verticalAlignment, replacementFormatting.verticalAlignment);
  assert.deepEqual(reimported.blocks[1]?.columnWidthsDxa, [2100, 4500, 2700]);
  const rendered = await verifyDocumentFile(outputPath, {
    outputDir: path.join(outputDir, "render"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(rendered.summary.verifyOk, true);
  assert.equal(rendered.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const clearedFormatting = { ...replacementFormatting };
  delete clearedFormatting.verticalAlignment;
  const clearPath = path.join(outputDir, "cleared-alignment.docx");
  const clearAuditPath = path.join(outputDir, "cleared-alignment.audit.json");
  const cleared = await editImportedTableFormatting({
    inputPath: outputPath,
    outputPath: clearPath,
    auditPath: clearAuditPath,
    tableBlockIndex: 1,
    expectedFormatting: replacementFormatting,
    replacementFormatting: clearedFormatting,
  });
  assert.deepEqual(cleared.audit.validation.changedParts, ["word/document.xml"]);
  const clearedBytes = await fs.readFile(clearPath);
  const clearedZip = await JSZip.loadAsync(clearedBytes);
  const clearedXml = await clearedZip.file("word/document.xml").async("text");
  assert.doesNotMatch(flatTables(clearedXml)[0], /<w:vAlign\b/);
  assert.equal((await DocumentFile.importDocx(await FileBlob.load(clearPath))).blocks[1]?.verticalAlignment, undefined);

  const restoredFormatting = { ...clearedFormatting, verticalAlignment: "top" };
  const restorePath = path.join(outputDir, "restored-alignment.docx");
  const restoreAuditPath = path.join(outputDir, "restored-alignment.audit.json");
  const restored = await editImportedTableFormatting({
    inputPath: clearPath,
    outputPath: restorePath,
    auditPath: restoreAuditPath,
    tableBlockIndex: 1,
    expectedFormatting: clearedFormatting,
    replacementFormatting: restoredFormatting,
  });
  assert.deepEqual(restored.audit.validation.changedParts, ["word/document.xml"]);
  const restoredZip = await JSZip.loadAsync(await fs.readFile(restorePath));
  assertFormattingMarkup(flatTables(await restoredZip.file("word/document.xml").async("text"))[0], restoredFormatting, 3, 3);
  assert.equal((await DocumentFile.importDocx(await FileBlob.load(restorePath))).blocks[1]?.verticalAlignment, "top");

  const cliOutput = path.join(outputDir, "cli-output.docx");
  const cliAudit = path.join(outputDir, "cli-audit.json");
  const cliReplacement = {
    indentDxa: 180,
    cellMarginsDxa: { top: 90, bottom: 110, start: 140, end: 160 },
    borderColor: "336699",
    borderSize: 0,
    headerFill: "EAF2F8",
    verticalAlignment: "top",
  };
  assert.deepEqual(parseTableFormattingEditCli([
    sourcePath, cliOutput, cliAudit, "1", JSON.stringify(sourceFormatting), JSON.stringify(cliReplacement),
  ]), {
    inputPath: sourcePath,
    outputPath: cliOutput,
    auditPath: cliAudit,
    tableBlockIndex: 1,
    expectedFormatting: sourceFormatting,
    replacementFormatting: cliReplacement,
  });
  const cli = spawnSync(process.execPath, [
    workflowPath, sourcePath, cliOutput, cliAudit, "1", JSON.stringify(sourceFormatting), JSON.stringify(cliReplacement),
  ], { cwd: repoRoot, encoding: "utf8" });
  assert.equal(cli.status, 0, cli.stderr);
  assert.deepEqual(JSON.parse(cli.stdout), {
    outputPath: cliOutput,
    auditPath: cliAudit,
    outputSha256: createHash("sha256").update(await fs.readFile(cliOutput)).digest("hex"),
    changedParts: ["word/document.xml"],
  });
  const cliImported = await DocumentFile.importDocx(await FileBlob.load(cliOutput));
  assert.equal(cliImported.blocks[1]?.headerFill, cliReplacement.headerFill);
  assert.equal(cliImported.blocks[1]?.verticalAlignment, cliReplacement.verticalAlignment);

  await assert.rejects(
    () => editImportedTableFormatting({
      inputPath: sourcePath,
      outputPath: path.join(outputDir, "mismatched.docx"),
      auditPath: path.join(outputDir, "mismatched.json"),
      tableBlockIndex: 1,
      expectedFormatting: { ...sourceFormatting, borderColor: "000000" },
      replacementFormatting,
    }),
    /direct formatting does not match the expected source value/,
  );
  await assert.rejects(
    () => editImportedTableFormatting({
      inputPath: sourcePath,
      outputPath: path.join(outputDir, "noop.docx"),
      auditPath: path.join(outputDir, "noop.json"),
      tableBlockIndex: 1,
      expectedFormatting: sourceFormatting,
      replacementFormatting: sourceFormatting,
    }),
    /replacementFormatting must differ from expectedFormatting/,
  );
  await assert.rejects(
    () => editImportedTableFormatting({
      inputPath: sourcePath,
      outputPath: path.join(outputDir, "invalid.docx"),
      auditPath: path.join(outputDir, "invalid.json"),
      tableBlockIndex: 1,
      expectedFormatting: sourceFormatting,
      replacementFormatting: { ...replacementFormatting, borderSize: 1 },
    }),
    /borderSize must be zero or from 2 through 96/,
  );
  await assert.rejects(
    () => editImportedTableFormatting({
      inputPath: sourcePath,
      outputPath: path.join(outputDir, "invalid-alignment.docx"),
      auditPath: path.join(outputDir, "invalid-alignment.json"),
      tableBlockIndex: 1,
      expectedFormatting: sourceFormatting,
      replacementFormatting: { ...replacementFormatting, verticalAlignment: "middle" },
    }),
    /verticalAlignment must be top, center, or bottom/,
  );
  const mixedAlignmentPath = path.join(outputDir, "mixed-alignment.docx");
  const mixedAlignmentZip = await JSZip.loadAsync(sourceBytes);
  let alignmentLeaves = 0;
  mixedAlignmentZip.file("word/document.xml", sourceXml.replace(/<w:vAlign\b[^>]*w:val="center"[^>]*\/>/g, (leaf) => {
    alignmentLeaves += 1;
    return alignmentLeaves === 2 ? leaf.replace('w:val="center"', 'w:val="bottom"') : leaf;
  }));
  assert.equal(alignmentLeaves, 9);
  await fs.writeFile(mixedAlignmentPath, await mixedAlignmentZip.generateAsync({ type: "nodebuffer" }), { flag: "wx" });
  await assert.rejects(
    () => editImportedTableFormatting({
      inputPath: mixedAlignmentPath,
      outputPath: path.join(outputDir, "mixed-alignment-output.docx"),
      auditPath: path.join(outputDir, "mixed-alignment.audit.json"),
      tableBlockIndex: 1,
      expectedFormatting: sourceFormatting,
      replacementFormatting,
    }),
    /direct formatting does not match the expected source value/,
  );
  const nonCanonicalPath = path.join(outputDir, "noncanonical.docx");
  const nonCanonicalZip = await JSZip.loadAsync(sourceBytes);
  nonCanonicalZip.file("word/document.xml", sourceXml.replace(/<w:tblInd w:w="120"/, '<w:tblInd w:w="0120"'));
  await fs.writeFile(nonCanonicalPath, await nonCanonicalZip.generateAsync({ type: "nodebuffer" }), { flag: "wx" });
  await assert.rejects(
    () => editImportedTableFormatting({
      inputPath: nonCanonicalPath,
      outputPath: path.join(outputDir, "noncanonical-output.docx"),
      auditPath: path.join(outputDir, "noncanonical-audit.json"),
      tableBlockIndex: 1,
      expectedFormatting: sourceFormatting,
      replacementFormatting,
    }),
    /canonical unsigned integer/,
  );
  assert.deepEqual(await fs.readFile(sourcePath), sourceBytes);
} finally {
  await fs.rm(outputDir, { recursive: true, force: true });
}

console.log("document table-formatting workflow smoke ok");
