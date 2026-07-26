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
const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-table-header-rows-workflow-"));
const nativeStatus = nativeDocumentRenderStatus();

function flatTables(xml) {
  return [...xml.matchAll(/<w:tbl\b[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
}

function rows(tableXml) {
  return [...tableXml.matchAll(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g)].map((match) => match[0]);
}

function headerRows(tableXml) {
  const values = rows(tableXml).map((row) => /<w:tblHeader\b[^>]*\/>/.test(row));
  const firstFalse = values.indexOf(false);
  assert.equal(values.slice(firstFalse < 0 ? values.length : firstFalse).some(Boolean), false, "native repeat headers must be a contiguous prefix");
  return values.filter(Boolean).length;
}

function rowFills(tableXml, rowIndex) {
  return [...(rows(tableXml)[rowIndex] || "").matchAll(/<w:shd\b[^>]*w:fill="([0-9A-F]+)"[^>]*\/>/g)].map((match) => match[1]);
}

try {
  const sourceDocument = DocumentModel.create({ name: "Source-bound table repeat-header edit", blocks: [] });
  sourceDocument.addParagraph("The first table is the source-bound repeat-header target.");
  sourceDocument.addTable({
    name: "target-table",
    values: [
      ["Quarter", "Revenue", "Margin"],
      ["Q1", "1.2M", "44%"],
      ["Q2", "1.4M", "46%"],
      ["Q3", "1.6M", "49%"],
    ],
    widthDxa: 9300,
    indentDxa: 120,
    columnWidthsDxa: [2100, 4500, 2700],
    cellMarginsDxa: { top: 80, bottom: 80, start: 120, end: 120 },
    borderColor: "445566",
    borderSize: 8,
    headerFill: "E2E8F0",
    headerRowCount: 1,
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
    headerRowCount: 1,
  });
  const sourcePath = path.join(outputDir, "source.docx");
  await (await DocumentFile.exportDocx(sourceDocument)).save(sourcePath);
  const sourceBytes = await fs.readFile(sourcePath);
  const sourceImported = await DocumentFile.importDocx(await FileBlob.load(sourcePath));
  assert.equal(sourceImported.blocks[1]?.sourceBound, true);
  assert.equal(sourceImported.blocks[1]?.headerRowCount, 1);
  assert.equal(sourceImported.blocks[3]?.headerRowCount, 1);

  const workflowPath = path.join(repoRoot, "skills", "documents", "skills", "documents", "examples", "officekit-table-header-rows-edit-workflow.mjs");
  const {
    editImportedTableHeaderRows,
    parseTableHeaderRowsEditCli,
    tableHeaderRowsCliOutput,
  } = await import(workflowPath);
  const outputPath = path.join(outputDir, "output.docx");
  const auditPath = path.join(outputDir, "audit.json");
  const result = await editImportedTableHeaderRows({
    inputPath: sourcePath,
    outputPath,
    auditPath,
    tableBlockIndex: 1,
    expectedHeaderRowCount: 1,
    replacementHeaderRowCount: 2,
  });
  assert.equal(result.audit.provider.actual, "office-kit");
  assert.equal(result.audit.provider.silentFallback, false);
  assert.deepEqual(result.audit.savePolicy, { strategy: "rewrite", noReplace: true });
  assert.equal(result.audit.operation.type, "source-bound-table-header-rows-edit");
  assert.deepEqual(result.audit.operation.target, { id: sourceImported.blocks[1].id, blockIndex: 1, tableOrdinal: 0 });
  assert.equal(result.audit.operation.expectedHeaderRowCount, 1);
  assert.equal(result.audit.operation.replacementHeaderRowCount, 2);
  assert.deepEqual(result.audit.validation.changedParts, ["word/document.xml"]);
  assert.equal(result.audit.validation.tableHeaderRowsXmlResidual.ok, true);
  assert.equal(result.audit.validation.reimport.headerRowCount, 2);
  assert.equal(result.audit.validation.nativeRenderRequired, true);
  assert.deepEqual(tableHeaderRowsCliOutput(result).changedParts, ["word/document.xml"]);
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
  assert.equal(headerRows(sourceTables[0]), 1);
  assert.equal(headerRows(outputTables[0]), 2);
  assert.equal(headerRows(sourceTables[1]), 1);
  assert.equal(headerRows(outputTables[1]), 1);
  assert.deepEqual(rowFills(sourceTables[0], 0), ["E2E8F0", "E2E8F0", "E2E8F0"]);
  assert.deepEqual(rowFills(outputTables[0], 0), ["E2E8F0", "E2E8F0", "E2E8F0"]);
  assert.deepEqual(rowFills(outputTables[0], 1), [], "repeat-header semantics must not turn headerFill into second-row styling");
  assert.equal(outputTables[1], sourceTables[1]);
  const reimported = await DocumentFile.importDocx(await FileBlob.load(outputPath));
  assert.equal(reimported.blocks[1]?.headerRowCount, 2);
  assert.equal(reimported.blocks[1]?.headerFill, "E2E8F0");
  assert.equal(reimported.blocks[3]?.headerRowCount, 1);
  const rendered = await verifyDocumentFile(outputPath, {
    outputDir: path.join(outputDir, "render"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(rendered.summary.verifyOk, true);
  assert.equal(rendered.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const cliOutput = path.join(outputDir, "cli-output.docx");
  const cliAudit = path.join(outputDir, "cli-audit.json");
  assert.deepEqual(parseTableHeaderRowsEditCli([sourcePath, cliOutput, cliAudit, "1", "1", "3"]), {
    inputPath: sourcePath,
    outputPath: cliOutput,
    auditPath: cliAudit,
    tableBlockIndex: 1,
    expectedHeaderRowCount: "1",
    replacementHeaderRowCount: "3",
  });
  const cli = spawnSync(process.execPath, [workflowPath, sourcePath, cliOutput, cliAudit, "1", "1", "3"], { cwd: repoRoot, encoding: "utf8" });
  assert.equal(cli.status, 0, cli.stderr);
  assert.deepEqual(JSON.parse(cli.stdout), {
    outputPath: cliOutput,
    auditPath: cliAudit,
    outputSha256: createHash("sha256").update(await fs.readFile(cliOutput)).digest("hex"),
    changedParts: ["word/document.xml"],
  });
  assert.equal((await DocumentFile.importDocx(await FileBlob.load(cliOutput))).blocks[1]?.headerRowCount, 3);

  await assert.rejects(
    () => editImportedTableHeaderRows({
      inputPath: sourcePath,
      outputPath: path.join(outputDir, "mismatched.docx"),
      auditPath: path.join(outputDir, "mismatched.json"),
      tableBlockIndex: 1,
      expectedHeaderRowCount: 0,
      replacementHeaderRowCount: 2,
    }),
    /repeat-header count does not match the expected source value/,
  );
  await assert.rejects(
    () => editImportedTableHeaderRows({
      inputPath: sourcePath,
      outputPath: path.join(outputDir, "noop.docx"),
      auditPath: path.join(outputDir, "noop.json"),
      tableBlockIndex: 1,
      expectedHeaderRowCount: 1,
      replacementHeaderRowCount: 1,
    }),
    /replacementHeaderRowCount must differ/,
  );
  await assert.rejects(
    () => editImportedTableHeaderRows({
      inputPath: sourcePath,
      outputPath: path.join(outputDir, "out-of-range.docx"),
      auditPath: path.join(outputDir, "out-of-range.json"),
      tableBlockIndex: 1,
      expectedHeaderRowCount: 1,
      replacementHeaderRowCount: 5,
    }),
    /integer from 0 through 4/,
  );
  const nonPrefixPath = path.join(outputDir, "non-prefix.docx");
  const nonPrefixZip = await JSZip.loadAsync(sourceBytes);
  let nonPrefixXml = sourceXml.replace("<w:trPr><w:tblHeader /></w:trPr>", "");
  const tableRows = [...nonPrefixXml.matchAll(/<w:tr>/g)];
  assert.ok(tableRows.length >= 2);
  const secondRowOffset = (tableRows[1].index ?? 0) + "<w:tr>".length;
  nonPrefixXml = `${nonPrefixXml.slice(0, secondRowOffset)}<w:trPr><w:tblHeader /></w:trPr>${nonPrefixXml.slice(secondRowOffset)}`;
  nonPrefixZip.file("word/document.xml", nonPrefixXml);
  await fs.writeFile(nonPrefixPath, await nonPrefixZip.generateAsync({ type: "nodebuffer" }), { flag: "wx" });
  await assert.rejects(
    () => editImportedTableHeaderRows({
      inputPath: nonPrefixPath,
      outputPath: path.join(outputDir, "non-prefix-output.docx"),
      auditPath: path.join(outputDir, "non-prefix-audit.json"),
      tableBlockIndex: 1,
      expectedHeaderRowCount: 0,
      replacementHeaderRowCount: 1,
    }),
    /non-prefix w:tblHeader row/,
  );
  assert.deepEqual(await fs.readFile(sourcePath), sourceBytes);
} finally {
  await fs.rm(outputDir, { recursive: true, force: true });
}

console.log("document table-header-rows workflow smoke ok");
