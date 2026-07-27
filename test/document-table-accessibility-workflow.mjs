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
const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-table-accessibility-workflow-"));
const nativeStatus = nativeDocumentRenderStatus();

function decodeXml(value) {
  return String(value).replace(/&(amp|lt|gt|quot|apos);/g, (_match, entity) => ({
    amp: "&", lt: "<", gt: ">", quot: "\"", apos: "'",
  })[entity]);
}

function flatTables(xml) {
  return [...String(xml).matchAll(/<w:tbl\b[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
}

function rawAccessibility(tableXml) {
  const metadata = {};
  for (const match of String(tableXml).matchAll(/<w:(tblCaption|tblDescription)\s+w:val="([^"]*)"\s*\/>/g)) {
    const field = match[1] === "tblCaption" ? "title" : "description";
    assert.equal(Object.hasOwn(metadata, field), false, `duplicate native ${match[1]} leaf`);
    metadata[field] = decodeXml(match[2]);
  }
  return metadata;
}

function tableProjection(document) {
  let tableOrdinal = 0;
  return document.blocks.flatMap((block, blockIndex) => {
    if (block.kind !== "table") return [];
    const snapshot = {
      id: block.id,
      blockIndex,
      tableOrdinal,
      values: block.values.map((row) => [...row]),
      cells: structuredClone(block.cells),
      widthDxa: block.widthDxa,
      indentDxa: block.indentDxa,
      columnWidthsDxa: [...block.columnWidthsDxa],
      headerRowCount: block.headerRowCount,
      keepTogetherRows: [...block.keepTogetherRows],
      minimumRowHeightsDxa: [...block.minimumRowHeightsDxa],
      accessibility: structuredClone(block.accessibility || {}),
    };
    tableOrdinal += 1;
    return [snapshot];
  });
}

try {
  const sourceDocument = DocumentModel.create({ name: "Source-bound table accessibility metadata edit", blocks: [] });
  sourceDocument.addParagraph("The first table has reviewed non-visible Word alternative text.");
  sourceDocument.addTable({
    name: "target-table",
    values: [["Gate", "Owner"], ["Accessibility review", "Release engineering"], ["Native render", "Document QA"]],
    widthDxa: 9300,
    indentDxa: 120,
    columnWidthsDxa: [3600, 5700],
    cellMarginsDxa: { top: 80, bottom: 80, start: 120, end: 120 },
    borderColor: "445566",
    borderSize: 8,
    headerFill: "E2E8F0",
    headerRowCount: 1,
    accessibility: {
      title: "Quarterly delivery readiness",
      description: "A two-column matrix covering release accessibility and native-render ownership.",
    },
  });
  sourceDocument.addParagraph("The second table is a source-bound metadata canary.");
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
    accessibility: {
      title: "Sibling table",
      description: "This alternative text must not change.",
    },
  });

  const sourcePath = path.join(outputDir, "source.docx");
  await (await DocumentFile.exportDocx(sourceDocument)).save(sourcePath);
  const sourceBytes = await fs.readFile(sourcePath);
  const sourceImported = await DocumentFile.importDocx(await FileBlob.load(sourcePath));
  const sourceTables = tableProjection(sourceImported);
  assert.deepEqual(sourceTables.map((table) => table.accessibility), [
    {
      title: "Quarterly delivery readiness",
      description: "A two-column matrix covering release accessibility and native-render ownership.",
    },
    { title: "Sibling table", description: "This alternative text must not change." },
  ]);

  const workflowPath = path.join(repoRoot, "skills", "documents", "skills", "documents", "examples", "officekit-table-accessibility-edit-workflow.mjs");
  const {
    editImportedTableAccessibility,
    parseTableAccessibilityEditCli,
    tableAccessibilityCliOutput,
  } = await import(workflowPath);
  const expectedAccessibility = sourceTables[0].accessibility;
  const replacementAccessibility = { title: "Release-readiness decision matrix" };
  const outputPath = path.join(outputDir, "output.docx");
  const auditPath = path.join(outputDir, "audit.json");
  const result = await editImportedTableAccessibility({
    inputPath: sourcePath,
    outputPath,
    auditPath,
    tableBlockIndex: 1,
    expectedAccessibility,
    replacementAccessibility,
  });
  assert.equal(result.audit.provider.actual, "office-kit");
  assert.equal(result.audit.provider.silentFallback, false);
  assert.deepEqual(result.audit.savePolicy, { strategy: "rewrite", noReplace: true });
  assert.equal(result.audit.operation.type, "source-bound-table-accessibility-edit");
  assert.deepEqual(result.audit.operation.target, { id: sourceImported.blocks[1].id, blockIndex: 1, tableOrdinal: 0 });
  assert.deepEqual(result.audit.operation.expectedAccessibility, expectedAccessibility);
  assert.deepEqual(result.audit.operation.replacementAccessibility, replacementAccessibility);
  assert.deepEqual(result.audit.validation.changedParts, ["word/document.xml"]);
  assert.equal(result.audit.validation.tableAccessibilityXmlResidual.ok, true);
  assert.deepEqual(result.audit.validation.reimport.accessibility, replacementAccessibility);
  assert.equal(result.audit.validation.nativeRenderRequired, true);
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
  const sourceNativeTables = flatTables(sourceXml);
  const outputNativeTables = flatTables(outputXml);
  assert.equal(sourceNativeTables.length, 2);
  assert.equal(outputNativeTables.length, 2);
  assert.deepEqual(rawAccessibility(sourceNativeTables[0]), expectedAccessibility);
  assert.deepEqual(rawAccessibility(outputNativeTables[0]), replacementAccessibility);
  assert.deepEqual(rawAccessibility(outputNativeTables[1]), sourceTables[1].accessibility);
  assert.equal(outputNativeTables[1], sourceNativeTables[1], "sibling table XML must remain byte-identical");

  const reimported = await DocumentFile.importDocx(await FileBlob.load(outputPath));
  const expectedTables = structuredClone(sourceTables);
  expectedTables[0].accessibility = replacementAccessibility;
  assert.deepEqual(tableProjection(reimported), expectedTables);
  assert.deepEqual(reimported.resolve(sourceImported.blocks[1].id)?.accessibility, replacementAccessibility);

  const baselineDir = path.join(outputDir, "visual-baseline");
  const sourceQa = await verifyDocumentFile(sourcePath, {
    outputDir: path.join(outputDir, "source-qa"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
    baselineDir,
    writeBaseline: true,
  });
  assert.equal(sourceQa.summary.verifyOk, true);
  const outputQa = await verifyDocumentFile(outputPath, {
    outputDir: path.join(outputDir, "output-qa"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
    baselineDir,
  });
  assert.equal(outputQa.summary.verifyOk, true);
  assert.equal(outputQa.summary.modelPixelDiff.changed, false, "non-visible table metadata must not affect the model render");
  assert.equal(outputQa.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  if (nativeStatus.available) {
    assert.equal(outputQa.summary.nativeRender.pageCountMatches, true);
    assert.equal(outputQa.summary.nativeRender.pages.every((page) => page.pixelDiff.changed === false), true, "non-visible table metadata must not affect native pages");
  }

  const cliOutput = path.join(outputDir, "cli-output.docx");
  const cliAudit = path.join(outputDir, "cli-audit.json");
  const cliReplacement = { title: "Final delivery readiness matrix", description: "Reviewed alternative text for the release gate matrix." };
  assert.deepEqual(parseTableAccessibilityEditCli([
    sourcePath, cliOutput, cliAudit, "1", JSON.stringify(expectedAccessibility), JSON.stringify(cliReplacement),
  ]), {
    inputPath: sourcePath,
    outputPath: cliOutput,
    auditPath: cliAudit,
    tableBlockIndex: 1,
    expectedAccessibility,
    replacementAccessibility: cliReplacement,
  });
  const cli = spawnSync(process.execPath, [
    workflowPath, sourcePath, cliOutput, cliAudit, "1", JSON.stringify(expectedAccessibility), JSON.stringify(cliReplacement),
  ], { cwd: repoRoot, encoding: "utf8" });
  assert.equal(cli.status, 0, cli.stderr);
  assert.deepEqual(JSON.parse(cli.stdout), {
    outputPath: cliOutput,
    auditPath: cliAudit,
    outputSha256: createHash("sha256").update(await fs.readFile(cliOutput)).digest("hex"),
    changedParts: ["word/document.xml"],
  });
  assert.deepEqual((await DocumentFile.importDocx(await FileBlob.load(cliOutput))).blocks[1]?.accessibility, cliReplacement);

  await assert.rejects(
    () => editImportedTableAccessibility({
      inputPath: sourcePath,
      outputPath: path.join(outputDir, "mismatched.docx"),
      auditPath: path.join(outputDir, "mismatched.json"),
      tableBlockIndex: 1,
      expectedAccessibility: { title: "Not the source title" },
      replacementAccessibility,
    }),
    /does not match the expected source value/,
  );
  await assert.rejects(
    () => editImportedTableAccessibility({
      inputPath: sourcePath,
      outputPath: path.join(outputDir, "noop.docx"),
      auditPath: path.join(outputDir, "noop.json"),
      tableBlockIndex: 1,
      expectedAccessibility,
      replacementAccessibility: expectedAccessibility,
    }),
    /must differ/,
  );
  assert.throws(
    () => parseTableAccessibilityEditCli([sourcePath, "x", "y", "1", "not-json", "{}"]),
    /must be a JSON object/,
  );
  assert.throws(
    () => parseTableAccessibilityEditCli([
      sourcePath, "x", "y", "1", JSON.stringify({ title: String.fromCharCode(0xd800) }), "{}",
    ]),
    /must contain 1 through 32767 XML-safe characters/,
  );

  const duplicatePath = path.join(outputDir, "duplicate-caption.docx");
  const duplicateZip = await JSZip.loadAsync(sourceBytes);
  duplicateZip.file(
    "word/document.xml",
    sourceXml.replace(
      '<w:tblCaption w:val="Quarterly delivery readiness" />',
      '<w:tblCaption w:val="Quarterly delivery readiness" /><w:tblCaption w:val="Conflicting title" />',
    ),
  );
  await fs.writeFile(duplicatePath, await duplicateZip.generateAsync({ type: "nodebuffer" }), { flag: "wx" });
  await assert.rejects(
    () => editImportedTableAccessibility({
      inputPath: duplicatePath,
      outputPath: path.join(outputDir, "duplicate-output.docx"),
      auditPath: path.join(outputDir, "duplicate-audit.json"),
      tableBlockIndex: 1,
      expectedAccessibility: {},
      replacementAccessibility: { title: "Requested replacement" },
    }),
    /duplicate w:tblCaption metadata/,
  );
  assert.deepEqual(await fs.readFile(sourcePath), sourceBytes);
} finally {
  await fs.rm(outputDir, { recursive: true, force: true });
}

console.log("document table-accessibility workflow smoke ok");
