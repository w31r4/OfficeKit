import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

import { DocumentFile, DocumentModel, FileBlob } from "office-kit";
import {
  createDocumentFromFixture,
  nativeDocumentRenderStatus,
  runDocumentFixture,
  verifyDocumentFile,
} from "./skill-harness/documents/scripts/workflow.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const fixturesDir = path.join(repoRoot, "test", "skill-harness", "documents", "fixtures");
const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-document-skill-"));
const baselineDir = path.join(outputDir, "baselines");
const nativeStatus = nativeDocumentRenderStatus();
const packagedRenderScript = path.join(repoRoot, "skills", "documents", "skills", "documents", "render_docx.py");
const packagedRenderSource = await fs.readFile(packagedRenderScript, "utf8");
assert.doesNotMatch(packagedRenderSource, /pdf2image/i, "the packaged renderer must not require an undeclared Python package");

async function runFixture(name, options = {}) {
  const result = await runDocumentFixture(path.join(fixturesDir, `${name}.json`), {
    outputDir: path.join(outputDir, name),
    nativeRender: "off",
    ...options,
  });
  assert.deepEqual(Object.keys(result).sort(), ["docxPath", "fixture", "qa"]);
  assert.equal(result.fixture.name, name);
  assert.equal(result.qa.summary.packageOk, true);
  assert.equal(result.qa.summary.verifyOk, true);
  assert.equal(result.qa.summary.visualQaOk, true);
  return result;
}

try {
  assert.equal(createDocumentFromFixture({ settings: { trackRevisions: true }, blocks: [] }).settings.trackRevisions, true);
  const business = await runFixture("business-brief", {
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(business.qa.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  if (nativeStatus.available) {
    assert.equal(business.qa.summary.nativeRender.ok, true);
    assert.ok(business.qa.summary.nativeRender.pageCount >= 1);
    assert.equal(business.qa.summary.nativeRender.pages.length, business.qa.summary.nativeRender.pageCount);
  }
  for (const filePath of Object.values(business.qa.summary.files)) {
    const stat = await fs.stat(filePath);
    assert.ok(stat.isFile() && stat.size > 0, `Expected non-empty document skill output ${filePath}`);
  }

  const renderPython = process.env.OFFICE_KIT_DOCUMENTS_PYTHON || "python3";
  const pythonVersion = spawnSync(renderPython, ["--version"], { encoding: "utf8" });
  if (nativeStatus.available && pythonVersion.status === 0) {
    const renderedOutput = path.join(outputDir, "packaged-render-docx");
    const rendered = spawnSync(renderPython, [
      packagedRenderScript,
      business.docxPath,
      "--output_dir",
      renderedOutput,
      "--emit_pdf",
    ], {
      cwd: repoRoot,
      encoding: "utf8",
      timeout: 60_000,
      env: { ...process.env, PYTHONNOUSERSITE: "1" },
    });
    assert.equal(rendered.status, 0, `The packaged renderer must work with a standard Python runtime\nSTDOUT:\n${rendered.stdout}\nSTDERR:\n${rendered.stderr}`);
    const renderedFiles = await fs.readdir(renderedOutput);
    assert.ok(renderedFiles.includes("page-1.png"), "the packaged renderer must promote canonical page PNG names");
    assert.ok(renderedFiles.includes(path.basename(business.docxPath, ".docx") + ".pdf"), "--emit_pdf must retain the native PDF for explicit QA");
    assert.ok((await fs.stat(path.join(renderedOutput, "page-1.png"))).size > 0);

    const pdfInfoProbe = spawnSync(renderPython, [
      "-c",
      "import importlib.util, sys; spec = importlib.util.spec_from_file_location('renderer', sys.argv[1]); renderer = importlib.util.module_from_spec(spec); spec.loader.exec_module(renderer); width, height = renderer._read_pdf_page_size(sys.argv[2], False); assert width > 0 and height > 0",
      packagedRenderScript,
      path.join(renderedOutput, path.basename(business.docxPath, ".docx") + ".pdf"),
    ], {
      cwd: repoRoot,
      encoding: "utf8",
      timeout: 30_000,
      env: { ...process.env, PYTHONNOUSERSITE: "1" },
    });
    assert.equal(pdfInfoProbe.status, 0, `The packaged renderer's Poppler pdfinfo fallback must work\nSTDOUT:\n${pdfInfoProbe.stdout}\nSTDERR:\n${pdfInfoProbe.stderr}`);

    if (!process.env.OFFICE_KIT_DOCUMENTS_PYTHON) {
      const pythonExecutable = spawnSync(renderPython, ["-c", "import sys; print(sys.executable)"], { encoding: "utf8" }).stdout.trim();
      const noPopplerPath = path.join(outputDir, "no-poppler-path");
      await fs.mkdir(noPopplerPath);
      const missingPoppler = spawnSync(pythonExecutable, [
        "-c",
        "import importlib.util, sys; spec = importlib.util.spec_from_file_location('renderer', sys.argv[1]); renderer = importlib.util.module_from_spec(spec); spec.loader.exec_module(renderer); renderer._read_pdf_page_size(sys.argv[2], False)",
        packagedRenderScript,
        path.join(renderedOutput, path.basename(business.docxPath, ".docx") + ".pdf"),
      ], {
        cwd: repoRoot,
        encoding: "utf8",
        timeout: 30_000,
        env: { ...process.env, PATH: noPopplerPath, PYTHONNOUSERSITE: "1" },
      });
      assert.notEqual(missingPoppler.status, 0, "a missing Poppler command must fail closed");
      assert.match(missingPoppler.stderr, /Poppler pdfinfo is unavailable: expected `pdfinfo` on PATH/i);
    }
  }

  const document = await DocumentFile.importDocx(await FileBlob.load(business.docxPath));
  assert.equal(document.defaultRunStyle.fontFamily, "Aptos");
  assert.equal(document.styles.values().some((style) => style.id === "BriefLead" && style.basedOn === "Normal"), true);
  const editedLead = document.blocks.find((block) => block.text === "Create, inspect, render, and verify the canonical DOCX path.");
  assert.equal(editedLead?.kind, "paragraph");
  assert.equal(editedLead?.paragraphFormat.alignment, "left");
  assert.equal(editedLead?.runs[0].style.color, "#9c2b2e");
  assert.equal(document.blocks.some((block) => block.kind === "listItem" && block.text === "Inspect stable document blocks and fields."), true);
  assert.equal(document.blocks.find((block) => block.kind === "table")?.values[1][2], "Pass");
  const hyperlink = document.blocks.find((block) => block.kind === "hyperlink");
  assert.equal(hyperlink?.url, "https://learn.microsoft.com/office/open-xml/word-processing");
  assert.equal(hyperlink?.tooltip, "Edited through the canonical Office path");
  assert.equal(hyperlink?.history, false);
  const image = document.blocks.find((block) => block.kind === "image");
  assert.equal(image?.alt, "Edited green status mark");
  assert.equal(image?.widthPx, 48);
  assert.equal(document.blocks.find((block) => block.kind === "field")?.instruction, "NUMPAGES");
  assert.equal(document.blocks.find((block) => block.kind === "section")?.margins.left, 1200);
  assert.equal(document.comments[0]?.author, "Lead reviewer");
  assert.equal(document.comments[0]?.text, "Delivery evidence approved.");
  assert.equal(document.settings.evenAndOddHeaders, true);
  assert.equal(document.headers.some((item) => item.referenceType === "first" && item.variantActive), true);
  assert.equal(document.footers.some((item) => item.referenceType === "even" && item.fieldInstruction === "PAGE"), true);

  const accessibilityAuditDir = path.join(outputDir, "accessibility-audit");
  const accessibilityReportPath = path.join(accessibilityAuditDir, "report.json");
  const businessSourceBeforeAudit = await fs.readFile(business.docxPath);
  const { auditDocxAccessibility } = await import(
    "../skills/documents/skills/documents/examples/officekit-accessibility-audit-workflow.mjs"
  );
  const accessibilityResult = await auditDocxAccessibility({
    inputPath: business.docxPath,
    reportPath: accessibilityReportPath,
    maxChars: 100_000,
  });
  assert.equal(accessibilityResult.report.schema, "office-kit.docx-accessibility-audit.v1");
  assert.equal(accessibilityResult.report.provider.requested, "office-kit");
  assert.equal(accessibilityResult.report.provider.actual, "office-kit");
  assert.equal(accessibilityResult.report.provider.silentFallback, false);
  assert.deepEqual(accessibilityResult.report.savePolicy, {
    strategy: "none",
    sourceMutation: false,
    artifactProduced: false,
  });
  assert.equal(accessibilityResult.report.operation.type, "document-accessibility-audit");
  assert.equal(accessibilityResult.report.accessibility.conformanceClaimed, false);
  assert.equal(accessibilityResult.report.accessibility.machineCheckPassed, false);
  assert.equal(accessibilityResult.report.accessibility.manualReviewRequired, true);
  assert.deepEqual(
    accessibilityResult.report.accessibility.issues.map((entry) => entry.type),
    ["tableHeaderRowMissing"],
  );
  assert.deepEqual(
    accessibilityResult.report.accessibility.manualChecks.map((entry) => entry.type),
    ["tablePurposeAndDescription"],
  );
  assert.equal(accessibilityResult.report.validation.sourceUnchanged, true);
  assert.equal(accessibilityResult.report.validation.documentVerify.ok, true);
  assert.equal(accessibilityResult.report.boundaries.tableAndLinkPurpose, "manual-author-review");
  assert.deepEqual(JSON.parse(await fs.readFile(accessibilityReportPath, "utf8")), accessibilityResult.report);
  assert.deepEqual(await fs.readFile(business.docxPath), businessSourceBeforeAudit);
  assert.deepEqual(await fs.readdir(accessibilityAuditDir), ["report.json"]);
  await assert.rejects(
    () => auditDocxAccessibility({ inputPath: business.docxPath, reportPath: accessibilityReportPath }),
    /already exists; refusing to overwrite/,
  );
  await assert.rejects(
    () => auditDocxAccessibility({ inputPath: business.docxPath, reportPath: business.docxPath }),
    /reportPath must be distinct from inputPath/,
  );

  const businessZip = await JSZip.loadAsync(await fs.readFile(business.docxPath));
  for (const part of [
    "word/document.xml",
    "word/styles.xml",
    "word/numbering.xml",
    "word/comments.xml",
    "word/settings.xml",
  ]) assert.ok(businessZip.file(part), `Expected ${part}`);
  assert.ok(Object.keys(businessZip.files).some((name) => /(^|\/)media\//.test(name)));
  assert.ok(Object.keys(businessZip.files).filter((name) => /^word\/header\d+\.xml$/.test(name)).length >= 2);
  assert.ok(Object.keys(businessZip.files).some((name) => /^word\/footer\d+\.xml$/.test(name)));
  const businessXml = await businessZip.file("word/document.xml").async("text");
  assert.match(businessXml, /w:instr="NUMPAGES"/);
  assert.match(businessXml, /<w:drawing>/);
  assert.match(businessXml, /<w:sectPr>/);
  assert.match(await fs.readFile(business.qa.summary.files.packageInspect, "utf8"), /word\/document\.xml/);

  const floating = await runFixture("office-kit-floating-image", {
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  const floatingDocument = await DocumentFile.importDocx(await FileBlob.load(floating.docxPath));
  const floatingImage = floatingDocument.blocks.find((block) => block.kind === "image");
  assert.deepEqual(floatingImage?.placement.horizontal, { relativeTo: "page", offsetPx: 150 });
  assert.deepEqual(floatingImage?.placement.vertical, { relativeTo: "paragraph", offsetPx: 0 });
  assert.equal(floatingImage?.placement.wrap, "topAndBottom");
  assert.equal(floatingImage?.placement.wrapSide, undefined);
  assert.deepEqual(floatingImage?.placement.distanceFromTextPx, { top: 8, right: 0, bottom: 8, left: 0 });
  assert.equal(floating.qa.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  if (nativeStatus.available) {
    assert.equal(floating.qa.summary.nativeRender.ok, true);
    assert.ok(floating.qa.summary.nativeRender.pages.length >= 1);
  }
  const floatingZip = await JSZip.loadAsync(await fs.readFile(floating.docxPath));
  const floatingXml = await floatingZip.file("word/document.xml").async("text");
  assert.match(floatingXml, /<wp:anchor(?=[^>]*behindDoc="0")(?=[^>]*allowOverlap="0")[^>]*>/);
  assert.match(floatingXml, /<wp:positionH relativeFrom="page"><wp:posOffset>1428750<\/wp:posOffset><\/wp:positionH>/);
  assert.match(floatingXml, /<wp:positionV relativeFrom="paragraph"><wp:posOffset>0<\/wp:posOffset><\/wp:positionV>/);
  assert.match(floatingXml, /<wp:wrapTopAndBottom\s*\/>/);

  const pictureBullets = await runFixture("office-kit-picture-bullets", {
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  const pictureBulletDocument = await DocumentFile.importDocx(await FileBlob.load(pictureBullets.docxPath));
  const pictureBulletItems = pictureBulletDocument.blocks.filter((block) => block.kind === "listItem" && block.pictureBullet);
  assert.equal(pictureBulletItems.length, 2);
  assert.deepEqual(pictureBulletItems[0].pictureBullet, pictureBulletItems[1].pictureBullet);
  assert.equal(pictureBulletItems[0].pictureBullet.widthPt, 16);
  assert.equal(pictureBulletItems[0].pictureBullet.heightPt, 14);
  assert.equal(pictureBulletItems[0].pictureBullet.alt, "Approved green action marker");
  assert.equal(pictureBullets.qa.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  if (nativeStatus.available) {
    assert.equal(pictureBullets.qa.summary.nativeRender.ok, true);
    assert.ok(pictureBullets.qa.summary.nativeRender.pages.length >= 1);
  }
  const pictureBulletZip = await JSZip.loadAsync(await fs.readFile(pictureBullets.docxPath));
  const pictureBulletNumberingXml = await pictureBulletZip.file("word/numbering.xml").async("text");
  assert.equal((pictureBulletNumberingXml.match(/<w:numPicBullet\b/g) || []).length, 2);
  assert.match(pictureBulletNumberingXml, /<w:lvlOverride[^>]*w:ilvl="0">[\s\S]*<w:lvlPicBulletId w:val="1"\s*\/>/);
  assert.match(pictureBulletNumberingXml, /<v:shape(?=[^>]*style="width:16pt;height:14pt")(?=[^>]*alt="Approved green action marker")[^>]*>/);

  const watermark = await runFixture("office-kit-watermark", {
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  const watermarkDocument = await DocumentFile.importDocx(await FileBlob.load(watermark.docxPath));
  assert.equal(watermarkDocument.watermarks.length, 1);
  assert.equal(watermarkDocument.watermarks[0].text, "INTERNAL REVIEW");
  assert.equal(watermarkDocument.watermarks[0].referenceType, "default");
  assert.equal(watermarkDocument.watermarks[0].sectionIndex, 0);
  assert.equal(watermarkDocument.watermarks[0].sourceBound, true);
  assert.equal(watermarkDocument.watermarks[0].editable, true);
  assert.equal(watermarkDocument.resolve(watermarkDocument.watermarks[0].id), watermarkDocument.watermarks[0]);
  assert.match(watermarkDocument.inspect({ kind: "watermark" }).ndjson, /"text":"INTERNAL REVIEW"/);
  assert.equal(watermark.qa.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  if (nativeStatus.available) {
    assert.equal(watermark.qa.summary.nativeRender.ok, true);
    assert.ok(watermark.qa.summary.nativeRender.pages.length >= 1);
  }
  const watermarkZip = await JSZip.loadAsync(await fs.readFile(watermark.docxPath));
  const watermarkHeaderPath = Object.keys(watermarkZip.files).find((name) => /^word\/header\d+\.xml$/.test(name));
  assert.ok(watermarkHeaderPath);
  assert.match(await watermarkZip.file(watermarkHeaderPath).async("text"), /<v:textpath[^>]*string="INTERNAL REVIEW"/);
  const watermarkSourceBytes = await fs.readFile(watermark.docxPath);
  const { editDocumentWatermark } = await import(
    "../skills/documents/skills/documents/examples/officekit-watermark-workflow.mjs"
  );
  const watermarkWorkflowOutput = path.join(outputDir, "watermark-approved.docx");
  const watermarkWorkflowAudit = path.join(outputDir, "watermark-approved-audit.json");
  const watermarkWorkflow = await editDocumentWatermark({
    inputPath: watermark.docxPath,
    outputPath: watermarkWorkflowOutput,
    auditPath: watermarkWorkflowAudit,
    expectedText: "INTERNAL REVIEW",
    replacementText: "APPROVED COPY",
    sectionIndex: 0,
    referenceType: "default",
  });
  assert.deepEqual(watermarkWorkflow.audit.operation.changedParts, [watermarkHeaderPath]);
  assert.equal(watermarkWorkflow.audit.provider.actual, "office-kit");
  assert.equal(watermarkWorkflow.audit.provider.silentFallback, false);
  assert.equal(watermarkWorkflow.audit.validation.secondImport, true);
  assert.equal(watermarkWorkflow.audit.validation.nativeRenderRequiredBeforeDelivery, true);
  assert.deepEqual(await fs.readFile(watermark.docxPath), watermarkSourceBytes);
  const watermarkWorkflowDocument = await DocumentFile.importDocx(await FileBlob.load(watermarkWorkflowOutput));
  assert.equal(watermarkWorkflowDocument.watermarks[0]?.text, "APPROVED COPY");

  const headerWorkflowSourcePath = path.join(outputDir, "header-text-source.docx");
  const headerWorkflowOutputPath = path.join(outputDir, "header-text-reviewed.docx");
  const headerWorkflowAuditPath = path.join(outputDir, "header-text-audit.json");
  const headerWorkflowDocument = DocumentModel.create({ name: "Header text workflow", blocks: [] });
  headerWorkflowDocument.addParagraph("Only the requested ordinary header text may change.");
  headerWorkflowDocument.addHeader("Northwind | Internal", { id: "header/review-target", sectionIndex: 0 });
  headerWorkflowDocument.addHeader("Retain the body and footer exactly.", { id: "header/companion", sectionIndex: 0 });
  headerWorkflowDocument.addFooter("1", { id: "footer/page", sectionIndex: 0, fieldInstruction: "PAGE" });
  await (await DocumentFile.exportDocx(headerWorkflowDocument)).save(headerWorkflowSourcePath);
  const headerWorkflowSourceBytes = await fs.readFile(headerWorkflowSourcePath);
  const { editImportedHeaderText } = await import(
    "../skills/documents/skills/documents/examples/officekit-header-text-edit-workflow.mjs"
  );
  const headerWorkflow = await editImportedHeaderText({
    inputPath: headerWorkflowSourcePath,
    outputPath: headerWorkflowOutputPath,
    auditPath: headerWorkflowAuditPath,
    expectedText: "Northwind | Internal",
    replacementText: "Northwind | Reviewed",
    sectionIndex: 0,
    referenceType: "default",
  });
  assert.equal(headerWorkflow.audit.provider.actual, "office-kit");
  assert.equal(headerWorkflow.audit.provider.silentFallback, false);
  assert.equal(headerWorkflow.audit.savePolicy.noReplace, true);
  assert.deepEqual(headerWorkflow.audit.validation.changedParts, ["word/header1.xml"]);
  assert.equal(headerWorkflow.audit.validation.reimport.editable, true);
  assert.deepEqual(await fs.readFile(headerWorkflowSourcePath), headerWorkflowSourceBytes);
  const headerWorkflowRoundTrip = await DocumentFile.importDocx(await FileBlob.load(headerWorkflowOutputPath));
  assert.equal(headerWorkflowRoundTrip.headers.find((header) => header.text === "Northwind | Reviewed")?.editable, true);
  assert.equal(headerWorkflowRoundTrip.headers.find((header) => header.text === "Retain the body and footer exactly.")?.partPath, "word/header1.xml");
  assert.equal(headerWorkflowRoundTrip.footers[0]?.fieldInstruction, "PAGE");
  const [headerWorkflowSourceZip, headerWorkflowOutputZip] = await Promise.all([
    JSZip.loadAsync(headerWorkflowSourceBytes),
    JSZip.loadAsync(await fs.readFile(headerWorkflowOutputPath)),
  ]);
  const headerSourceXml = await headerWorkflowSourceZip.file("word/header1.xml").async("text");
  const headerOutputXml = await headerWorkflowOutputZip.file("word/header1.xml").async("text");
  assert.equal(headerSourceXml.replace("Northwind | Internal", "__target__"), headerOutputXml.replace("Northwind | Reviewed", "__target__"));
  for (const name of Object.keys(headerWorkflowSourceZip.files).filter((entry) => !headerWorkflowSourceZip.files[entry].dir && entry !== "word/header1.xml")) {
    assert.deepEqual(
      Buffer.from(await headerWorkflowSourceZip.file(name).async("uint8array")),
      Buffer.from(await headerWorkflowOutputZip.file(name).async("uint8array")),
      `Unexpected source-bound header workflow drift in ${name}`,
    );
  }
  await assert.rejects(
    () => editImportedHeaderText({
      inputPath: headerWorkflowSourcePath,
      outputPath: headerWorkflowOutputPath,
      auditPath: path.join(outputDir, "header-text-second-audit.json"),
      expectedText: "Northwind | Internal",
      replacementText: "Northwind | Reviewed",
    }),
    /outputPath already exists/i,
  );

  const footerWorkflowSourcePath = path.join(outputDir, "footer-text-source.docx");
  const footerWorkflowOutputPath = path.join(outputDir, "footer-text-reviewed.docx");
  const footerWorkflowAuditPath = path.join(outputDir, "footer-text-audit.json");
  const footerWorkflowDocument = DocumentModel.create({ name: "Footer text workflow", blocks: [] });
  footerWorkflowDocument.addParagraph("Only the requested ordinary footer text may change.");
  footerWorkflowDocument.addHeader("1", { id: "header/page", sectionIndex: 0, fieldInstruction: "PAGE" });
  footerWorkflowDocument.addFooter("Northwind | Internal", { id: "footer/review-target", sectionIndex: 0 });
  footerWorkflowDocument.addFooter("Retain the header and body exactly.", { id: "footer/companion", sectionIndex: 0 });
  await (await DocumentFile.exportDocx(footerWorkflowDocument)).save(footerWorkflowSourcePath);
  const footerWorkflowSourceBytes = await fs.readFile(footerWorkflowSourcePath);
  const { editImportedFooterText } = await import(
    "../skills/documents/skills/documents/examples/officekit-footer-text-edit-workflow.mjs"
  );
  const footerWorkflow = await editImportedFooterText({
    inputPath: footerWorkflowSourcePath,
    outputPath: footerWorkflowOutputPath,
    auditPath: footerWorkflowAuditPath,
    expectedText: "Northwind | Internal",
    replacementText: "Northwind | Reviewed",
    sectionIndex: 0,
    referenceType: "default",
  });
  assert.equal(footerWorkflow.audit.provider.actual, "office-kit");
  assert.equal(footerWorkflow.audit.provider.silentFallback, false);
  assert.equal(footerWorkflow.audit.savePolicy.noReplace, true);
  assert.deepEqual(footerWorkflow.audit.validation.changedParts, ["word/footer1.xml"]);
  assert.equal(footerWorkflow.audit.validation.reimport.editable, true);
  assert.deepEqual(await fs.readFile(footerWorkflowSourcePath), footerWorkflowSourceBytes);
  const footerWorkflowRoundTrip = await DocumentFile.importDocx(await FileBlob.load(footerWorkflowOutputPath));
  assert.equal(footerWorkflowRoundTrip.footers.find((footer) => footer.text === "Northwind | Reviewed")?.editable, true);
  assert.equal(footerWorkflowRoundTrip.footers.find((footer) => footer.text === "Retain the header and body exactly.")?.partPath, "word/footer1.xml");
  assert.equal(footerWorkflowRoundTrip.headers[0]?.fieldInstruction, "PAGE");
  const [footerWorkflowSourceZip, footerWorkflowOutputZip] = await Promise.all([
    JSZip.loadAsync(footerWorkflowSourceBytes),
    JSZip.loadAsync(await fs.readFile(footerWorkflowOutputPath)),
  ]);
  const footerSourceXml = await footerWorkflowSourceZip.file("word/footer1.xml").async("text");
  const footerOutputXml = await footerWorkflowOutputZip.file("word/footer1.xml").async("text");
  assert.equal(footerSourceXml.replace("Northwind | Internal", "__target__"), footerOutputXml.replace("Northwind | Reviewed", "__target__"));
  for (const name of Object.keys(footerWorkflowSourceZip.files).filter((entry) => !footerWorkflowSourceZip.files[entry].dir && entry !== "word/footer1.xml")) {
    assert.deepEqual(
      Buffer.from(await footerWorkflowSourceZip.file(name).async("uint8array")),
      Buffer.from(await footerWorkflowOutputZip.file(name).async("uint8array")),
      `Unexpected source-bound footer workflow drift in ${name}`,
    );
  }
  const watermarkRemovalOutput = path.join(outputDir, "watermark-removed.docx");
  const watermarkRemovalAudit = path.join(outputDir, "watermark-removed-audit.json");
  const watermarkRemoval = await editDocumentWatermark({
    inputPath: watermarkWorkflowOutput,
    outputPath: watermarkRemovalOutput,
    auditPath: watermarkRemovalAudit,
    expectedText: "APPROVED COPY",
    sectionIndex: 0,
    referenceType: "default",
    remove: true,
  });
  assert.equal(watermarkRemoval.audit.operation.type, "canonical-text-watermark-remove");
  assert.deepEqual(watermarkRemoval.audit.operation.changedParts, [watermarkHeaderPath]);
  assert.equal((await DocumentFile.importDocx(await FileBlob.load(watermarkRemovalOutput))).watermarks.length, 0);

  const merged = await runFixture("office-kit-merged-table");
  const mergedDocument = await DocumentFile.importDocx(await FileBlob.load(merged.docxPath));
  const mergedTable = mergedDocument.blocks.find((block) => block.kind === "table");
  assert.equal(mergedTable?.values[0][0], "Edited merged owner");
  assert.equal(mergedTable?.getCell(0, 0).columnSpan, 2);
  assert.equal(mergedTable?.getCell(0, 0).rowSpan, 2);
  assert.equal(mergedTable?.getCell(1, 0).verticalMerge, "continue");
  assert.equal(mergedTable?.getCell(1, 0).editable, false);
  assert.equal(mergedTable?.widthDxa, 9300);
  assert.deepEqual(mergedTable?.columnWidthsDxa, [2500, 3100, 3700]);
  assert.equal(mergedTable?.borderColor, "884400");

  const numbering = await runFixture("office-kit-numbering-edit");
  const numberingDocument = await DocumentFile.importDocx(await FileBlob.load(numbering.docxPath));
  const numberedItems = numberingDocument.blocks.filter((block) => block.kind === "listItem");
  assert.equal(numberedItems.length, 2);
  assert.equal(numberedItems[0].text, "Edited first grouped item");
  assert.equal(numberedItems.every((block) => block.numberFormat === "lowerRoman" && block.start === 5 && block.levelText === "%1."), true);

  const comments = await runFixture("office-kit-comments");
  const commentsDocument = await DocumentFile.importDocx(await FileBlob.load(comments.docxPath));
  assert.equal(commentsDocument.comments.length, 1);
  assert.equal(commentsDocument.comments[0].author, "Lead reviewer");
  assert.equal(commentsDocument.comments[0].initials, "LR");
  assert.equal(commentsDocument.comments[0].text, "Approved after source-bound review.");

  const controls = await runFixture("office-kit-content-controls", {
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  const controlsDocument = await DocumentFile.importDocx(await FileBlob.load(controls.docxPath));
  assert.deepEqual(controlsDocument.contentControls.map((control) => [control.tag, control.controlType, control.controlType === "checkbox" ? control.checked : control.text]), [
    ["CUSTOMER_NAME", "text", "Grace Hopper"],
    ["ACCOUNT_ID", "text", "AC-2048"],
    ["PRIORITY", "dropdown", "High"],
    ["CONTACT_METHOD", "comboBox", "Pager duty"],
    ["REVIEW_DATE", "date", "2028-02-29"],
    ["APPROVED", "checkbox", true],
    ["EXECUTIVE_SUMMARY", "text", "Ready for approval"],
    ["TABLE_OWNER", "text", "Katherine Johnson"],
    ["TABLE_APPROVED", "checkbox", true],
    ["TABLE_PRIORITY", "dropdown", "High"],
    ["TABLE_CONTACT", "comboBox", "In person"],
    ["TABLE_REVIEW_DATE", "date", "2028-02-29"],
  ]);
  assert.equal(controls.qa.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  assert.equal(controlsDocument.inspect({ kind: "contentControl" }).ndjson.includes("Customer name"), true);
  const controlsZip = await JSZip.loadAsync(await fs.readFile(controls.docxPath));
  const controlsXml = await controlsZip.file("word/document.xml").async("text");
  assert.equal((controlsXml.match(/<w:sdt>/g) || []).length, 12);
  assert.match(controlsXml, /<w:tag w:val="CUSTOMER_NAME"\s*\/>/);
  assert.match(controlsXml, /<w:tag w:val="ACCOUNT_ID"\s*\/>/);
  assert.match(controlsXml, /<w:tag w:val="APPROVED"\s*\/>/);
  assert.match(controlsXml, /<w14:checkbox>[\s\S]*<w14:checked w14:val="1"\s*\/>/);
  assert.match(controlsXml, /<w:tag w:val="PRIORITY"\s*\/>[\s\S]*<w:dropDownList w:lastValue="high">[\s\S]*<w:listItem(?=[^>]*w:displayText="High")(?=[^>]*w:value="high")[^>]*\/>/);
  assert.match(controlsXml, /<w:tag w:val="CONTACT_METHOD"\s*\/>[\s\S]*<w:comboBox w:lastValue="Pager duty">[\s\S]*<w:listItem(?=[^>]*w:displayText="Phone call")(?=[^>]*w:value="phone")[^>]*\/>/);
  assert.match(controlsXml, /<w:tag w:val="REVIEW_DATE"\s*\/>[\s\S]*<w:date w:fullDate="2028-02-29T00:00:00Z">[\s\S]*<w:dateFormat w:val="yyyy-MM-dd"\s*\/>/);
  assert.match(controlsXml, /<w:body>[\s\S]*<w:sdt>[\s\S]*?<w:tag w:val="EXECUTIVE_SUMMARY"\s*\/>[\s\S]*?<w:text\s*\/>[\s\S]*?<w:sdtContent>\s*<w:p>[\s\S]*Ready for approval[\s\S]*?<\/w:p>\s*<\/w:sdtContent>\s*<\/w:sdt>/);
  assert.equal(controlsDocument.contentControls.find((control) => control.tag === "EXECUTIVE_SUMMARY").placement, "block");
  const tableCellControl = controlsDocument.contentControls.find((control) => control.tag === "TABLE_OWNER");
  assert.equal(tableCellControl.placement, "tableCell");
  assert.equal(tableCellControl.text, "Katherine Johnson");
  assert.match(controlsXml, /<w:tc>[\s\S]*?<w:sdt>[\s\S]*?<w:tag w:val="TABLE_OWNER"\s*\/>[\s\S]*?Katherine Johnson[\s\S]*?<\/w:sdt>[\s\S]*?<\/w:tc>/);
  for (const tag of ["TABLE_APPROVED", "TABLE_PRIORITY", "TABLE_CONTACT", "TABLE_REVIEW_DATE"]) {
    assert.equal(controlsDocument.contentControls.find((control) => control.tag === tag).placement, "tableCell");
  }
  assert.match(controlsXml, /<w:tc>[\s\S]*?<w:tag w:val="TABLE_APPROVED"\s*\/>[\s\S]*?<w14:checkbox>[\s\S]*?<w14:checked w14:val="1"\s*\/>[\s\S]*?<\/w:sdt>[\s\S]*?<\/w:tc>/);
  assert.match(controlsXml, /<w:tc>[\s\S]*?<w:tag w:val="TABLE_PRIORITY"\s*\/>[\s\S]*?<w:dropDownList w:lastValue="high">[\s\S]*?<\/w:sdt>[\s\S]*?<\/w:tc>/);
  assert.match(controlsXml, /<w:tc>[\s\S]*?<w:tag w:val="TABLE_CONTACT"\s*\/>[\s\S]*?<w:comboBox w:lastValue="In person">[\s\S]*?<\/w:sdt>[\s\S]*?<\/w:tc>/);
  assert.match(controlsXml, /<w:tc>[\s\S]*?<w:tag w:val="TABLE_REVIEW_DATE"\s*\/>[\s\S]*?<w:date w:fullDate="2028-02-29T00:00:00Z">[\s\S]*?<\/w:sdt>[\s\S]*?<\/w:tc>/);

  const bibliography = await runFixture("office-kit-bibliography");
  const bibliographyDocument = await DocumentFile.importDocx(await FileBlob.load(bibliography.docxPath));
  assert.equal(bibliographyDocument.bibliography.styleName, "APA");
  assert.deepEqual(bibliographyDocument.bibliographySources.map((source) => [source.tag, source.title, source.authors[0].first]), [
    ["AgentSource", "Notes on the Analytical Engine", "Augusta Ada"],
  ]);
  assert.equal(bibliographyDocument.blocks.find((block) => block.kind === "citation")?.text, "(Lovelace, 1843, revised)");
  assert.equal(bibliographyDocument.settings.updateFields, true);
  assert.equal(bibliographyDocument.blocks.find((block) => block.kind === "field")?.instruction, "BIBLIOGRAPHY");
  assert.equal(bibliographyDocument.blocks.find((block) => block.kind === "field")?.display, "Refresh bibliography before delivery");
  const bibliographyZip = await JSZip.loadAsync(await fs.readFile(bibliography.docxPath));
  const bibliographyParts = Object.keys(bibliographyZip.files).filter((name) => /^customXml\/item\d*\.xml$/.test(name));
  assert.equal(bibliographyParts.length, 1);
  assert.match(await bibliographyZip.file(bibliographyParts[0]).async("text"), /<Sources\b[^>]*xmlns="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/bibliography"/);
  assert.match(await bibliographyZip.file("word/document.xml").async("text"), /w:instr=" CITATION AgentSource "/);
  assert.match(await bibliographyZip.file("word/document.xml").async("text"), /w:instr="BIBLIOGRAPHY"/);

  const multiParagraphNotes = await runFixture("office-kit-multi-paragraph-notes", {
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  const multiParagraphNotesDocument = await DocumentFile.importDocx(await FileBlob.load(multiParagraphNotes.docxPath));
  assert.deepEqual(multiParagraphNotesDocument.notes.map((note) => [note.kind, note.paragraphs]), [
    ["footnote", ["The integration test passed against the bundled codec.", "Native page render evidence was reviewed before delivery."]],
    ["endnote", ["OfficeKit preserves the source-bound anchor and note identity.", "Exactly two physical plain-text paragraphs remain editable."]],
  ]);
  assert.equal(multiParagraphNotes.qa.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  const multiParagraphNotesZip = await JSZip.loadAsync(await fs.readFile(multiParagraphNotes.docxPath));
  const footnoteXml = await multiParagraphNotesZip.file("word/footnotes.xml").async("text");
  const endnoteXml = await multiParagraphNotesZip.file("word/endnotes.xml").async("text");
  assert.match(footnoteXml, /The integration test passed against the bundled codec\./);
  assert.match(footnoteXml, /Native page render evidence was reviewed before delivery\./);
  assert.match(endnoteXml, /OfficeKit preserves the source-bound anchor and note identity\./);
  assert.match(endnoteXml, /Exactly two physical plain-text paragraphs remain editable\./);

  const noteWorkflowSourceDocument = DocumentModel.create({ name: "Source-bound note text edit", blocks: [] });
  const noteWorkflowFootnoteTarget = noteWorkflowSourceDocument.addParagraph("The release decision has a source-bound footnote.");
  const noteWorkflowEndnoteTarget = noteWorkflowSourceDocument.addParagraph("The architecture decision has a source-bound endnote.");
  noteWorkflowSourceDocument.addParagraph("This ordinary body paragraph must remain unchanged.");
  noteWorkflowSourceDocument.addFootnote(noteWorkflowFootnoteTarget, undefined, {
    paragraphs: ["Pilot report, section 4.2.", "The independent review & release audit are retained."],
  });
  noteWorkflowSourceDocument.addEndnote(noteWorkflowEndnoteTarget, "Architecture decision record 11.");
  const noteWorkflowSourcePath = path.join(outputDir, "source-bound-notes-source.docx");
  await (await DocumentFile.exportDocx(noteWorkflowSourceDocument)).save(noteWorkflowSourcePath);
  const noteWorkflowSourceBytes = await fs.readFile(noteWorkflowSourcePath);
  const noteWorkflowImported = await DocumentFile.importDocx(await FileBlob.load(noteWorkflowSourcePath));
  const noteWorkflowFootnote = noteWorkflowImported.notes.find((note) => note.kind === "footnote");
  const noteWorkflowEndnote = noteWorkflowImported.notes.find((note) => note.kind === "endnote");
  assert.ok(noteWorkflowFootnote && noteWorkflowEndnote);
  assert.deepEqual(noteWorkflowFootnote.paragraphs, ["Pilot report, section 4.2.", "The independent review & release audit are retained."]);
  const {
    editImportedNoteParagraphText,
    noteTextCliOutput,
    parseNoteTextEditCli,
  } = await import(
    "../skills/documents/skills/documents/examples/officekit-note-text-edit-workflow.mjs"
  );
  const footnoteTarget = {
    kind: "footnote",
    noteId: noteWorkflowFootnote.id,
    nativeId: noteWorkflowFootnote.nativeId,
    targetId: noteWorkflowFootnote.targetId,
    paragraphIndex: 0,
    expectedText: "Pilot report, section 4.2.",
  };
  const noteWorkflowOutputPath = path.join(outputDir, "source-bound-footnote-edited.docx");
  const noteWorkflowAuditPath = path.join(outputDir, "source-bound-footnote-edited-audit.json");
  const noteWorkflow = await editImportedNoteParagraphText({
    inputPath: noteWorkflowSourcePath,
    outputPath: noteWorkflowOutputPath,
    auditPath: noteWorkflowAuditPath,
    target: footnoteTarget,
    replacementText: "Pilot report, section 4.2, independently reviewed.",
  });
  assert.equal(noteWorkflow.audit.provider.actual, "office-kit");
  assert.equal(noteWorkflow.audit.provider.silentFallback, false);
  assert.deepEqual(noteWorkflow.audit.savePolicy, { strategy: "rewrite", noReplace: true });
  assert.deepEqual(noteWorkflow.audit.validation.changedParts, ["word/footnotes.xml"]);
  assert.equal(noteWorkflow.audit.validation.noteXmlResidual.ok, true);
  assert.equal(noteWorkflow.audit.validation.noteXmlResidual.partPath, "word/footnotes.xml");
  assert.equal(noteWorkflow.audit.validation.reimport.noteId, noteWorkflowFootnote.id);
  assert.equal(noteWorkflow.audit.validation.reimport.nativeId, noteWorkflowFootnote.nativeId);
  assert.equal(noteWorkflow.audit.validation.nativeRenderRequired, true);
  assert.deepEqual(noteTextCliOutput(noteWorkflow).changedParts, ["word/footnotes.xml"]);
  assert.deepEqual(await fs.readFile(noteWorkflowSourcePath), noteWorkflowSourceBytes);
  const noteWorkflowOutputBytes = await fs.readFile(noteWorkflowOutputPath);
  const [noteWorkflowSourceZip, noteWorkflowOutputZip] = await Promise.all([
    JSZip.loadAsync(noteWorkflowSourceBytes),
    JSZip.loadAsync(noteWorkflowOutputBytes),
  ]);
  const noteWorkflowPartPaths = Object.keys(noteWorkflowSourceZip.files).filter((partPath) => !noteWorkflowSourceZip.files[partPath].dir).sort();
  assert.deepEqual(
    Object.keys(noteWorkflowOutputZip.files).filter((partPath) => !noteWorkflowOutputZip.files[partPath].dir).sort(),
    noteWorkflowPartPaths,
  );
  for (const partPath of noteWorkflowPartPaths) {
    if (partPath === "word/footnotes.xml") continue;
    assert.deepEqual(
      Buffer.from(await noteWorkflowOutputZip.file(partPath).async("uint8array")),
      Buffer.from(await noteWorkflowSourceZip.file(partPath).async("uint8array")),
      `Only word/footnotes.xml may change; ${partPath} drifted.`,
    );
  }
  const noteWorkflowFootnoteXml = await noteWorkflowOutputZip.file("word/footnotes.xml").async("text");
  assert.match(noteWorkflowFootnoteXml, /<w:footnote w:id="1"><w:p><w:r><w:footnoteRef\s*\/><\/w:r><w:r><w:t xml:space="preserve"> Pilot report, section 4\.2, independently reviewed\.<\/w:t>/);
  assert.match(noteWorkflowFootnoteXml, /The independent review &amp; release audit are retained\./);
  const noteWorkflowOutput = await DocumentFile.importDocx(await FileBlob.load(noteWorkflowOutputPath));
  assert.deepEqual(noteWorkflowOutput.notes.map((note) => [note.kind, note.nativeId, note.targetId, note.paragraphs]), [
    ["footnote", noteWorkflowFootnote.nativeId, noteWorkflowFootnote.targetId, ["Pilot report, section 4.2, independently reviewed.", "The independent review & release audit are retained."]],
    ["endnote", noteWorkflowEndnote.nativeId, noteWorkflowEndnote.targetId, ["Architecture decision record 11."]],
  ]);
  assert.equal(noteWorkflowOutput.blocks[2].text, "This ordinary body paragraph must remain unchanged.");
  const noteWorkflowRender = await verifyDocumentFile(noteWorkflowOutputPath, {
    outputDir: path.join(outputDir, "source-bound-footnote-render"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(noteWorkflowRender.summary.verifyOk, true);
  assert.equal(noteWorkflowRender.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const endnoteSourceBytes = await fs.readFile(noteWorkflowOutputPath);
  const endnoteInput = await DocumentFile.importDocx(await FileBlob.load(noteWorkflowOutputPath));
  const endnoteTarget = endnoteInput.notes.find((note) => note.kind === "endnote");
  const noteWorkflowEndnoteOutputPath = path.join(outputDir, "source-bound-endnote-edited.docx");
  const noteWorkflowEndnoteAuditPath = path.join(outputDir, "source-bound-endnote-edited-audit.json");
  const endnoteWorkflow = await editImportedNoteParagraphText({
    inputPath: noteWorkflowOutputPath,
    outputPath: noteWorkflowEndnoteOutputPath,
    auditPath: noteWorkflowEndnoteAuditPath,
    target: {
      kind: "endnote",
      noteId: endnoteTarget.id,
      nativeId: endnoteTarget.nativeId,
      targetId: endnoteTarget.targetId,
      paragraphIndex: 0,
      expectedText: "Architecture decision record 11.",
    },
    replacementText: "Architecture decision record 11, approved for release.",
  });
  assert.deepEqual(endnoteWorkflow.audit.validation.changedParts, ["word/endnotes.xml"]);
  assert.deepEqual(await fs.readFile(noteWorkflowOutputPath), endnoteSourceBytes);
  const [endnoteSourceZip, endnoteOutputZip] = await Promise.all([
    JSZip.loadAsync(endnoteSourceBytes),
    JSZip.loadAsync(await fs.readFile(noteWorkflowEndnoteOutputPath)),
  ]);
  for (const partPath of Object.keys(endnoteSourceZip.files).filter((name) => !endnoteSourceZip.files[name].dir)) {
    if (partPath === "word/endnotes.xml") continue;
    assert.deepEqual(
      Buffer.from(await endnoteOutputZip.file(partPath).async("uint8array")),
      Buffer.from(await endnoteSourceZip.file(partPath).async("uint8array")),
      `Only word/endnotes.xml may change; ${partPath} drifted.`,
    );
  }

  const noteCliInput = await DocumentFile.importDocx(await FileBlob.load(noteWorkflowOutputPath));
  const noteCliFootnote = noteCliInput.notes.find((note) => note.kind === "footnote");
  const noteCliOutputPath = path.join(outputDir, "source-bound-note-cli.docx");
  const noteCliAuditPath = path.join(outputDir, "source-bound-note-cli-audit.json");
  const noteCliTarget = {
    kind: "footnote",
    noteId: noteCliFootnote.id,
    nativeId: noteCliFootnote.nativeId,
    targetId: noteCliFootnote.targetId,
    paragraphIndex: 1,
    expectedText: "The independent review & release audit are retained.",
  };
  assert.deepEqual(parseNoteTextEditCli([
    noteWorkflowOutputPath,
    noteCliOutputPath,
    noteCliAuditPath,
    JSON.stringify(noteCliTarget),
    "The independent review & approval audit are retained.",
  ]), {
    inputPath: noteWorkflowOutputPath,
    outputPath: noteCliOutputPath,
    auditPath: noteCliAuditPath,
    target: noteCliTarget,
    replacementText: "The independent review & approval audit are retained.",
  });
  const noteCliProcess = spawnSync(process.execPath, [
    path.join(repoRoot, "skills", "documents", "skills", "documents", "examples", "officekit-note-text-edit-workflow.mjs"),
    noteWorkflowOutputPath,
    noteCliOutputPath,
    noteCliAuditPath,
    JSON.stringify(noteCliTarget),
    "The independent review & approval audit are retained.",
  ], { cwd: repoRoot, encoding: "utf8" });
  assert.equal(noteCliProcess.status, 0, noteCliProcess.stderr);
  assert.deepEqual(JSON.parse(noteCliProcess.stdout), {
    outputPath: noteCliOutputPath,
    auditPath: noteCliAuditPath,
    outputSha256: createHash("sha256").update(await fs.readFile(noteCliOutputPath)).digest("hex"),
    changedParts: ["word/footnotes.xml"],
  });
  const noteCliDocument = await DocumentFile.importDocx(await FileBlob.load(noteCliOutputPath));
  assert.equal(noteCliDocument.notes.find((note) => note.kind === "footnote").paragraphs[1], "The independent review & approval audit are retained.");

  await assert.rejects(
    () => editImportedNoteParagraphText({
      inputPath: noteWorkflowSourcePath,
      outputPath: path.join(outputDir, "source-bound-note-mismatched.docx"),
      auditPath: path.join(outputDir, "source-bound-note-mismatched-audit.json"),
      target: { ...footnoteTarget, expectedText: "Wrong source precondition." },
      replacementText: "Never publish this value.",
    }),
    /does not match the expected source text/,
  );
  await assert.rejects(
    () => editImportedNoteParagraphText({
      inputPath: noteWorkflowSourcePath,
      outputPath: path.join(outputDir, "source-bound-note-wrong-native-id.docx"),
      auditPath: path.join(outputDir, "source-bound-note-wrong-native-id-audit.json"),
      target: { ...footnoteTarget, nativeId: footnoteTarget.nativeId + 1 },
      replacementText: "Never publish this value.",
    }),
    /Expected exactly one footnote matching the inspected note ID, native ID, and target ID; found 0/,
  );
  await assert.rejects(
    () => editImportedNoteParagraphText({
      inputPath: noteWorkflowSourcePath,
      outputPath: path.join(outputDir, "source-bound-note-line-break.docx"),
      auditPath: path.join(outputDir, "source-bound-note-line-break-audit.json"),
      target: footnoteTarget,
      replacementText: "This must\nfail closed.",
    }),
    /one XML-safe physical paragraph without line breaks/,
  );
  await assert.rejects(
    () => editImportedNoteParagraphText({
      inputPath: noteWorkflowSourcePath,
      outputPath: noteWorkflowSourcePath,
      auditPath: path.join(outputDir, "source-bound-note-overwrite-audit.json"),
      target: footnoteTarget,
      replacementText: "Never overwrite the source.",
    }),
    /must be distinct/,
  );

  const marginWorkflowSourceDocument = DocumentModel.create({ name: "Source-bound section margin edit", blocks: [] });
  marginWorkflowSourceDocument.addParagraph("Prelude for the canonical section-margin transaction.");
  marginWorkflowSourceDocument.addSection({
    breakType: "nextPage",
    margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, gutter: 0 },
    pageNumbering: { start: 1, format: "lowerRoman" },
  });
  marginWorkflowSourceDocument.addParagraph("Only this section's left page margin may change.");
  marginWorkflowSourceDocument.addSection({
    breakType: "nextPage",
    margins: { top: 720, right: 1440, bottom: 1440, left: 1440, gutter: 0 },
    pageNumbering: { start: 1, format: "decimal" },
  });
  marginWorkflowSourceDocument.addParagraph("The sibling section is a raw-XML canary.");
  for (let sectionIndex = 0; sectionIndex < 3; sectionIndex += 1) {
    marginWorkflowSourceDocument.addFooter("1", {
      id: `footer/margin-canary-${sectionIndex}`,
      sectionIndex,
      referenceType: "default",
      fieldInstruction: "PAGE",
    });
  }
  const marginWorkflowSourcePath = path.join(outputDir, "source-bound-section-margins-source.docx");
  await (await DocumentFile.exportDocx(marginWorkflowSourceDocument)).save(marginWorkflowSourcePath);
  const marginWorkflowSourceBytes = await fs.readFile(marginWorkflowSourcePath);
  const marginWorkflowImported = await DocumentFile.importDocx(await FileBlob.load(marginWorkflowSourcePath));
  assert.deepEqual(marginWorkflowImported.blocks[1]?.margins, { top: 1440, right: 1440, bottom: 1440, left: 1440, gutter: 0 });
  assert.deepEqual(marginWorkflowImported.blocks[3]?.margins, { top: 720, right: 1440, bottom: 1440, left: 1440, gutter: 0 });
  assert.deepEqual(marginWorkflowImported.blocks[1]?.pageNumbering, { start: 1, format: "lowerRoman" });
  assert.deepEqual(marginWorkflowImported.blocks[3]?.pageNumbering, { start: 1, format: "decimal" });
  const { editImportedSectionPageNumbering } = await import(
    "../skills/documents/skills/documents/examples/officekit-section-page-numbering-edit-workflow.mjs"
  );
  const pageNumberingRegressionOutputPath = path.join(outputDir, "source-bound-section-page-numbering-regression.docx");
  const pageNumberingRegressionAuditPath = path.join(outputDir, "source-bound-section-page-numbering-regression-audit.json");
  const pageNumberingRegression = await editImportedSectionPageNumbering({
    inputPath: marginWorkflowSourcePath,
    outputPath: pageNumberingRegressionOutputPath,
    auditPath: pageNumberingRegressionAuditPath,
    sectionBlockIndex: 1,
    expectedPageNumbering: { start: 1, format: "lowerRoman" },
    replacementPageNumbering: { start: 1, format: "upperRoman" },
  });
  assert.deepEqual(pageNumberingRegression.audit.validation.changedParts, ["word/document.xml"]);
  assert.deepEqual(await fs.readFile(marginWorkflowSourcePath), marginWorkflowSourceBytes);
  const pageNumberingRegressionDocument = await DocumentFile.importDocx(await FileBlob.load(pageNumberingRegressionOutputPath));
  assert.deepEqual(pageNumberingRegressionDocument.blocks[1]?.pageNumbering, { start: 1, format: "upperRoman" });
  assert.deepEqual(pageNumberingRegressionDocument.blocks[3]?.pageNumbering, { start: 1, format: "decimal" });
  const {
    editImportedSectionMargins,
    parseSectionMarginEditCli,
    sectionMarginCliOutput,
  } = await import(
    "../skills/documents/skills/documents/examples/officekit-section-margin-edit-workflow.mjs"
  );
  const sourceMargins = { top: 1440, right: 1440, bottom: 1440, left: 1440, gutter: 0 };
  const replacementMargins = { top: 1440, right: 1440, bottom: 1440, left: 1728, gutter: 0 };
  const marginWorkflowOutputPath = path.join(outputDir, "source-bound-section-margins-edited.docx");
  const marginWorkflowAuditPath = path.join(outputDir, "source-bound-section-margins-edited-audit.json");
  const marginWorkflow = await editImportedSectionMargins({
    inputPath: marginWorkflowSourcePath,
    outputPath: marginWorkflowOutputPath,
    auditPath: marginWorkflowAuditPath,
    sectionBlockIndex: 1,
    expectedMargins: sourceMargins,
    replacementMargins,
  });
  assert.equal(marginWorkflow.audit.provider.actual, "office-kit");
  assert.equal(marginWorkflow.audit.provider.silentFallback, false);
  assert.deepEqual(marginWorkflow.audit.savePolicy, { strategy: "rewrite", noReplace: true });
  assert.equal(marginWorkflow.audit.operation.type, "source-bound-section-margin-edit");
  assert.deepEqual(marginWorkflow.audit.operation.target, {
    id: marginWorkflowImported.blocks[1].id,
    blockIndex: 1,
    sectionOrdinal: 0,
  });
  assert.deepEqual(marginWorkflow.audit.operation.sourceMargins, sourceMargins);
  assert.deepEqual(marginWorkflow.audit.operation.replacementMargins, replacementMargins);
  assert.deepEqual(marginWorkflow.audit.validation.changedParts, ["word/document.xml"]);
  assert.equal(marginWorkflow.audit.validation.marginsXmlResidual.ok, true);
  assert.equal(marginWorkflow.audit.validation.marginsXmlResidual.headerTwips, 720);
  assert.equal(marginWorkflow.audit.validation.marginsXmlResidual.footerTwips, 720);
  assert.equal(marginWorkflow.audit.validation.reimport.editable, true);
  assert.deepEqual(marginWorkflow.audit.validation.reimport.margins, replacementMargins);
  assert.equal(marginWorkflow.audit.validation.nativeRenderRequired, true);
  assert.deepEqual(sectionMarginCliOutput(marginWorkflow).changedParts, ["word/document.xml"]);
  assert.deepEqual(await fs.readFile(marginWorkflowSourcePath), marginWorkflowSourceBytes);

  const marginWorkflowOutputBytes = await fs.readFile(marginWorkflowOutputPath);
  const [marginWorkflowSourceZip, marginWorkflowOutputZip] = await Promise.all([
    JSZip.loadAsync(marginWorkflowSourceBytes),
    JSZip.loadAsync(marginWorkflowOutputBytes),
  ]);
  const marginWorkflowParts = Object.keys(marginWorkflowSourceZip.files).filter((partPath) => !marginWorkflowSourceZip.files[partPath].dir).sort();
  assert.deepEqual(
    Object.keys(marginWorkflowOutputZip.files).filter((partPath) => !marginWorkflowOutputZip.files[partPath].dir).sort(),
    marginWorkflowParts,
  );
  for (const partPath of marginWorkflowParts) {
    if (partPath === "word/document.xml") continue;
    assert.deepEqual(
      Buffer.from(await marginWorkflowOutputZip.file(partPath).async("uint8array")),
      Buffer.from(await marginWorkflowSourceZip.file(partPath).async("uint8array")),
      `Only word/document.xml may change; ${partPath} drifted.`,
    );
  }
  const [marginSourceXml, marginOutputXml] = await Promise.all([
    marginWorkflowSourceZip.file("word/document.xml").async("text"),
    marginWorkflowOutputZip.file("word/document.xml").async("text"),
  ]);
  const pageMarginTags = (xml) => [...xml.matchAll(/<w:pgMar\b[^>]*\/>/g)].map((match) => match[0]);
  const pageMarginAttributes = (tag) => Object.fromEntries(
    [...tag.matchAll(/w:([\w-]+)="([^"]*)"/g)].map((match) => [match[1], Number(match[2])]),
  );
  const sourcePageMargins = pageMarginTags(marginSourceXml);
  const outputPageMargins = pageMarginTags(marginOutputXml);
  assert.equal(sourcePageMargins.length, 3);
  assert.equal(outputPageMargins.length, sourcePageMargins.length);
  assert.deepEqual(pageMarginAttributes(sourcePageMargins[0]), {
    top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0,
  });
  assert.deepEqual(pageMarginAttributes(outputPageMargins[0]), {
    top: 1440, right: 1440, bottom: 1440, left: 1728, header: 720, footer: 720, gutter: 0,
  });
  assert.deepEqual(outputPageMargins.slice(1), sourcePageMargins.slice(1));
  const marginWorkflowOutputDocument = await DocumentFile.importDocx(await FileBlob.load(marginWorkflowOutputPath));
  assert.deepEqual(marginWorkflowOutputDocument.blocks[1]?.margins, replacementMargins);
  assert.deepEqual(marginWorkflowOutputDocument.blocks[3]?.margins, { top: 720, right: 1440, bottom: 1440, left: 1440, gutter: 0 });
  const marginWorkflowRender = await verifyDocumentFile(marginWorkflowOutputPath, {
    outputDir: path.join(outputDir, "source-bound-section-margins-render"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(marginWorkflowRender.summary.verifyOk, true);
  assert.equal(marginWorkflowRender.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const marginCliOutputPath = path.join(outputDir, "source-bound-section-margins-cli.docx");
  const marginCliAuditPath = path.join(outputDir, "source-bound-section-margins-cli-audit.json");
  const marginCliReplacement = { top: 1440, right: 1440, bottom: 1440, left: 1584, gutter: 0 };
  assert.deepEqual(parseSectionMarginEditCli([
    marginWorkflowSourcePath,
    marginCliOutputPath,
    marginCliAuditPath,
    "1",
    JSON.stringify(sourceMargins),
    JSON.stringify(marginCliReplacement),
  ]), {
    inputPath: marginWorkflowSourcePath,
    outputPath: marginCliOutputPath,
    auditPath: marginCliAuditPath,
    sectionBlockIndex: 1,
    expectedMargins: sourceMargins,
    replacementMargins: marginCliReplacement,
  });
  const marginCliProcess = spawnSync(process.execPath, [
    path.join(repoRoot, "skills", "documents", "skills", "documents", "examples", "officekit-section-margin-edit-workflow.mjs"),
    marginWorkflowSourcePath,
    marginCliOutputPath,
    marginCliAuditPath,
    "1",
    JSON.stringify(sourceMargins),
    JSON.stringify(marginCliReplacement),
  ], { cwd: repoRoot, encoding: "utf8" });
  assert.equal(marginCliProcess.status, 0, marginCliProcess.stderr);
  assert.deepEqual(JSON.parse(marginCliProcess.stdout), {
    outputPath: marginCliOutputPath,
    auditPath: marginCliAuditPath,
    outputSha256: createHash("sha256").update(await fs.readFile(marginCliOutputPath)).digest("hex"),
    changedParts: ["word/document.xml"],
  });
  const marginCliDocument = await DocumentFile.importDocx(await FileBlob.load(marginCliOutputPath));
  assert.deepEqual(marginCliDocument.blocks[1]?.margins, marginCliReplacement);

  const {
    editImportedSectionPageGeometry,
    parseSectionPageGeometryEditCli,
    sectionPageGeometryCliOutput,
  } = await import(
    "../skills/documents/skills/documents/examples/officekit-section-page-geometry-edit-workflow.mjs"
  );
  const { wordAttributes } = await import(
    "../skills/documents/skills/documents/artifact_tool/_source_bound_sections.mjs"
  );
  const sourcePageGeometry = {
    orientation: "portrait",
    pageSize: { widthTwips: 12240, heightTwips: 15840 },
  };
  const replacementPageGeometry = {
    orientation: "landscape",
    pageSize: { widthTwips: 15840, heightTwips: 12240 },
  };
  const pageGeometryWorkflowOutputPath = path.join(outputDir, "source-bound-section-page-geometry-edited.docx");
  const pageGeometryWorkflowAuditPath = path.join(outputDir, "source-bound-section-page-geometry-edited-audit.json");
  const pageGeometryWorkflow = await editImportedSectionPageGeometry({
    inputPath: marginWorkflowSourcePath,
    outputPath: pageGeometryWorkflowOutputPath,
    auditPath: pageGeometryWorkflowAuditPath,
    sectionBlockIndex: 1,
    expectedPageGeometry: sourcePageGeometry,
    replacementPageGeometry,
  });
  assert.equal(pageGeometryWorkflow.audit.provider.actual, "office-kit");
  assert.equal(pageGeometryWorkflow.audit.provider.silentFallback, false);
  assert.deepEqual(pageGeometryWorkflow.audit.savePolicy, { strategy: "rewrite", noReplace: true });
  assert.equal(pageGeometryWorkflow.audit.operation.type, "source-bound-section-page-geometry-edit");
  assert.deepEqual(pageGeometryWorkflow.audit.operation.target, {
    id: marginWorkflowImported.blocks[1].id,
    blockIndex: 1,
    sectionOrdinal: 0,
  });
  assert.deepEqual(pageGeometryWorkflow.audit.operation.sourcePageGeometry, sourcePageGeometry);
  assert.deepEqual(pageGeometryWorkflow.audit.operation.replacementPageGeometry, replacementPageGeometry);
  assert.deepEqual(pageGeometryWorkflow.audit.validation.changedParts, ["word/document.xml"]);
  assert.equal(pageGeometryWorkflow.audit.validation.pageSizeXmlResidual.ok, true);
  assert.deepEqual(pageGeometryWorkflow.audit.validation.reimport.pageGeometry, replacementPageGeometry);
  assert.equal(pageGeometryWorkflow.audit.validation.reimport.editable, true);
  assert.equal(pageGeometryWorkflow.audit.validation.nativeRenderRequired, true);
  assert.deepEqual(sectionPageGeometryCliOutput(pageGeometryWorkflow).changedParts, ["word/document.xml"]);
  assert.deepEqual(await fs.readFile(marginWorkflowSourcePath), marginWorkflowSourceBytes);

  const pageGeometryWorkflowOutputBytes = await fs.readFile(pageGeometryWorkflowOutputPath);
  const [pageGeometrySourceZip, pageGeometryOutputZip] = await Promise.all([
    JSZip.loadAsync(marginWorkflowSourceBytes),
    JSZip.loadAsync(pageGeometryWorkflowOutputBytes),
  ]);
  assert.deepEqual(
    Object.keys(pageGeometryOutputZip.files).filter((partPath) => !pageGeometryOutputZip.files[partPath].dir).sort(),
    marginWorkflowParts,
  );
  for (const partPath of marginWorkflowParts) {
    if (partPath === "word/document.xml") continue;
    assert.deepEqual(
      Buffer.from(await pageGeometryOutputZip.file(partPath).async("uint8array")),
      Buffer.from(await pageGeometrySourceZip.file(partPath).async("uint8array")),
      `Only word/document.xml may change; ${partPath} drifted.`,
    );
  }
  const [pageGeometrySourceXml, pageGeometryOutputXml] = await Promise.all([
    pageGeometrySourceZip.file("word/document.xml").async("text"),
    pageGeometryOutputZip.file("word/document.xml").async("text"),
  ]);
  const pageSizeTags = (xml) => [...xml.matchAll(/<w:pgSz\b[^>]*\/>/g)].map((match) => match[0]);
  const pageSizeAttributes = (tag) => Object.fromEntries(
    [...tag.matchAll(/w:([\w-]+)="([^"]*)"/g)].map((match) => [match[1], match[1] === "orient" ? match[2] : Number(match[2])]),
  );
  const sourcePageSizes = pageSizeTags(pageGeometrySourceXml);
  const outputPageSizes = pageSizeTags(pageGeometryOutputXml);
  assert.equal(sourcePageSizes.length, 3);
  assert.equal(outputPageSizes.length, sourcePageSizes.length);
  assert.deepEqual(pageSizeAttributes(sourcePageSizes[0]), { w: 12240, h: 15840, orient: "portrait" });
  assert.deepEqual(pageSizeAttributes(outputPageSizes[0]), { w: 15840, h: 12240, orient: "landscape" });
  assert.deepEqual(outputPageSizes.slice(1), sourcePageSizes.slice(1));
  const pageGeometryWorkflowOutputDocument = await DocumentFile.importDocx(await FileBlob.load(pageGeometryWorkflowOutputPath));
  assert.equal(pageGeometryWorkflowOutputDocument.blocks[1]?.orientation, "landscape");
  assert.deepEqual(pageGeometryWorkflowOutputDocument.blocks[1]?.pageSize, replacementPageGeometry.pageSize);
  assert.equal(pageGeometryWorkflowOutputDocument.blocks[3]?.orientation, "portrait");
  assert.deepEqual(pageGeometryWorkflowOutputDocument.blocks[3]?.pageSize, sourcePageGeometry.pageSize);
  const pageGeometryWorkflowRender = await verifyDocumentFile(pageGeometryWorkflowOutputPath, {
    outputDir: path.join(outputDir, "source-bound-section-page-geometry-render"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(pageGeometryWorkflowRender.summary.verifyOk, true);
  assert.equal(pageGeometryWorkflowRender.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const pageGeometryCliOutputPath = path.join(outputDir, "source-bound-section-page-geometry-cli.docx");
  const pageGeometryCliAuditPath = path.join(outputDir, "source-bound-section-page-geometry-cli-audit.json");
  const pageGeometryCliReplacement = {
    orientation: "landscape",
    pageSize: { widthTwips: 15120, heightTwips: 12240 },
  };
  assert.deepEqual(parseSectionPageGeometryEditCli([
    marginWorkflowSourcePath,
    pageGeometryCliOutputPath,
    pageGeometryCliAuditPath,
    "1",
    JSON.stringify(sourcePageGeometry),
    JSON.stringify(pageGeometryCliReplacement),
  ]), {
    inputPath: marginWorkflowSourcePath,
    outputPath: pageGeometryCliOutputPath,
    auditPath: pageGeometryCliAuditPath,
    sectionBlockIndex: 1,
    expectedPageGeometry: sourcePageGeometry,
    replacementPageGeometry: pageGeometryCliReplacement,
  });
  const pageGeometryCliProcess = spawnSync(process.execPath, [
    path.join(repoRoot, "skills", "documents", "skills", "documents", "examples", "officekit-section-page-geometry-edit-workflow.mjs"),
    marginWorkflowSourcePath,
    pageGeometryCliOutputPath,
    pageGeometryCliAuditPath,
    "1",
    JSON.stringify(sourcePageGeometry),
    JSON.stringify(pageGeometryCliReplacement),
  ], { cwd: repoRoot, encoding: "utf8" });
  assert.equal(pageGeometryCliProcess.status, 0, pageGeometryCliProcess.stderr);
  assert.deepEqual(JSON.parse(pageGeometryCliProcess.stdout), {
    outputPath: pageGeometryCliOutputPath,
    auditPath: pageGeometryCliAuditPath,
    outputSha256: createHash("sha256").update(await fs.readFile(pageGeometryCliOutputPath)).digest("hex"),
    changedParts: ["word/document.xml"],
  });
  const pageGeometryCliDocument = await DocumentFile.importDocx(await FileBlob.load(pageGeometryCliOutputPath));
  assert.equal(pageGeometryCliDocument.blocks[1]?.orientation, "landscape");
  assert.deepEqual(pageGeometryCliDocument.blocks[1]?.pageSize, pageGeometryCliReplacement.pageSize);

  assert.throws(
    () => wordAttributes('<w:pgSz w:w="12240" x:w="1" w:h="15840" w:orient="portrait"/>', "adversarial w:pgSz"),
    /noncanonical attribute namespace: x:w/,
  );
  assert.throws(
    () => wordAttributes('<w:pgSz w:w="12240" w:w="1" w:h="15840" w:orient="portrait"/>', "adversarial w:pgSz"),
    /duplicate or invalid w: attribute: w:w/,
  );
  assert.throws(
    () => wordAttributes("<w:pgSz w:w='12240' w:h=\"15840\" w:orient=\"portrait\"/>", "adversarial w:pgSz"),
    /unsupported XML attribute syntax/,
  );
  await assert.rejects(
    () => editImportedSectionPageGeometry({
      inputPath: marginWorkflowSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-page-geometry-mismatched.docx"),
      auditPath: path.join(outputDir, "source-bound-section-page-geometry-mismatched-audit.json"),
      sectionBlockIndex: 1,
      expectedPageGeometry: { ...sourcePageGeometry, pageSize: { ...sourcePageGeometry.pageSize, widthTwips: 999 } },
      replacementPageGeometry,
    }),
    /page geometry does not match the expected source value/,
  );
  await assert.rejects(
    () => editImportedSectionPageGeometry({
      inputPath: marginWorkflowSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-page-geometry-invalid.docx"),
      auditPath: path.join(outputDir, "source-bound-section-page-geometry-invalid-audit.json"),
      sectionBlockIndex: 1,
      expectedPageGeometry: sourcePageGeometry,
      replacementPageGeometry: { orientation: "landscape", pageSize: { widthTwips: 15840 } },
    }),
    /replacementPageGeometry\.pageSize\.heightTwips is required/,
  );
  await assert.rejects(
    () => editImportedSectionPageGeometry({
      inputPath: marginWorkflowSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-page-geometry-noop.docx"),
      auditPath: path.join(outputDir, "source-bound-section-page-geometry-noop-audit.json"),
      sectionBlockIndex: 1,
      expectedPageGeometry: sourcePageGeometry,
      replacementPageGeometry: sourcePageGeometry,
    }),
    /replacementPageGeometry must differ from expectedPageGeometry/,
  );
  const irregularPageGeometrySourcePath = path.join(outputDir, "source-bound-section-page-geometry-irregular-source.docx");
  const irregularPageGeometryZip = await JSZip.loadAsync(marginWorkflowSourceBytes);
  const irregularPageGeometryXml = await irregularPageGeometryZip.file("word/document.xml").async("text");
  irregularPageGeometryZip.file("word/document.xml", irregularPageGeometryXml.replace(/<w:pgSz\b[^>]*\/>/, (tag) => tag.replace(/\/>$/, " w:code=\"1\"/>")));
  await fs.writeFile(irregularPageGeometrySourcePath, await irregularPageGeometryZip.generateAsync({ type: "nodebuffer" }), { flag: "wx" });
  await assert.rejects(
    () => editImportedSectionPageGeometry({
      inputPath: irregularPageGeometrySourcePath,
      outputPath: path.join(outputDir, "source-bound-section-page-geometry-irregular.docx"),
      auditPath: path.join(outputDir, "source-bound-section-page-geometry-irregular-audit.json"),
      sectionBlockIndex: 1,
      expectedPageGeometry: sourcePageGeometry,
      replacementPageGeometry,
    }),
    /unsupported w:pgSz attributes: code/,
  );
  assert.deepEqual(await fs.readFile(marginWorkflowSourcePath), marginWorkflowSourceBytes);

  const {
    editImportedSectionLineNumbering,
    parseSectionLineNumberingEditCli,
    sectionLineNumberingCliOutput,
  } = await import(
    "../skills/documents/skills/documents/examples/officekit-section-line-numbering-edit-workflow.mjs"
  );
  const lineNumberingSourceDocument = DocumentModel.create({ name: "Source-bound section line-numbering edit", blocks: [] });
  lineNumberingSourceDocument.addParagraph("Prelude for the canonical section line-numbering transaction.");
  lineNumberingSourceDocument.addSection({
    breakType: "nextPage",
    lineNumbering: { countBy: 5, start: 0, distance: 360, restart: "newPage" },
  });
  lineNumberingSourceDocument.addParagraph("Only this section's line-number cadence, offset, distance, and restart behavior may change.");
  lineNumberingSourceDocument.addSection({
    breakType: "nextPage",
    lineNumbering: { countBy: 2, start: 3, distance: 240, restart: "newSection" },
  });
  lineNumberingSourceDocument.addParagraph("The sibling section is a raw-XML canary.");
  const lineNumberingSourcePath = path.join(outputDir, "source-bound-section-line-numbering-source.docx");
  await (await DocumentFile.exportDocx(lineNumberingSourceDocument)).save(lineNumberingSourcePath);
  const lineNumberingSourceBytes = await fs.readFile(lineNumberingSourcePath);
  const lineNumberingImported = await DocumentFile.importDocx(await FileBlob.load(lineNumberingSourcePath));
  const sourceLineNumbering = { countBy: 5, start: 0, distance: 360, restart: "newPage" };
  const replacementLineNumbering = { countBy: 10, start: 4, distance: 480, restart: "continuous" };
  assert.deepEqual(lineNumberingImported.blocks[1]?.lineNumbering, sourceLineNumbering);
  assert.deepEqual(lineNumberingImported.blocks[3]?.lineNumbering, { countBy: 2, start: 3, distance: 240, restart: "newSection" });

  const lineNumberingWorkflowOutputPath = path.join(outputDir, "source-bound-section-line-numbering-edited.docx");
  const lineNumberingWorkflowAuditPath = path.join(outputDir, "source-bound-section-line-numbering-edited-audit.json");
  const lineNumberingWorkflow = await editImportedSectionLineNumbering({
    inputPath: lineNumberingSourcePath,
    outputPath: lineNumberingWorkflowOutputPath,
    auditPath: lineNumberingWorkflowAuditPath,
    sectionBlockIndex: 1,
    expectedLineNumbering: sourceLineNumbering,
    replacementLineNumbering,
  });
  assert.equal(lineNumberingWorkflow.audit.provider.actual, "office-kit");
  assert.equal(lineNumberingWorkflow.audit.provider.silentFallback, false);
  assert.deepEqual(lineNumberingWorkflow.audit.savePolicy, { strategy: "rewrite", noReplace: true });
  assert.equal(lineNumberingWorkflow.audit.operation.type, "source-bound-section-line-numbering-edit");
  assert.deepEqual(lineNumberingWorkflow.audit.operation.target, {
    id: lineNumberingImported.blocks[1].id,
    blockIndex: 1,
    sectionOrdinal: 0,
  });
  assert.deepEqual(lineNumberingWorkflow.audit.operation.sourceLineNumbering, sourceLineNumbering);
  assert.deepEqual(lineNumberingWorkflow.audit.operation.replacementLineNumbering, replacementLineNumbering);
  assert.deepEqual(lineNumberingWorkflow.audit.validation.changedParts, ["word/document.xml"]);
  assert.equal(lineNumberingWorkflow.audit.validation.lineNumberingXmlResidual.ok, true);
  assert.deepEqual(lineNumberingWorkflow.audit.validation.reimport.lineNumbering, replacementLineNumbering);
  assert.equal(lineNumberingWorkflow.audit.validation.reimport.editable, true);
  assert.equal(lineNumberingWorkflow.audit.validation.nativeRenderRequired, true);
  assert.deepEqual(sectionLineNumberingCliOutput(lineNumberingWorkflow).changedParts, ["word/document.xml"]);
  assert.deepEqual(await fs.readFile(lineNumberingSourcePath), lineNumberingSourceBytes);

  const lineNumberingWorkflowOutputBytes = await fs.readFile(lineNumberingWorkflowOutputPath);
  const [lineNumberingSourceZip, lineNumberingOutputZip] = await Promise.all([
    JSZip.loadAsync(lineNumberingSourceBytes),
    JSZip.loadAsync(lineNumberingWorkflowOutputBytes),
  ]);
  const lineNumberingParts = Object.keys(lineNumberingSourceZip.files).filter((partPath) => !lineNumberingSourceZip.files[partPath].dir).sort();
  assert.deepEqual(
    Object.keys(lineNumberingOutputZip.files).filter((partPath) => !lineNumberingOutputZip.files[partPath].dir).sort(),
    lineNumberingParts,
  );
  for (const partPath of lineNumberingParts) {
    if (partPath === "word/document.xml") continue;
    assert.deepEqual(
      Buffer.from(await lineNumberingOutputZip.file(partPath).async("uint8array")),
      Buffer.from(await lineNumberingSourceZip.file(partPath).async("uint8array")),
      `Only word/document.xml may change; ${partPath} drifted.`,
    );
  }
  const [lineNumberingSourceXml, lineNumberingOutputXml] = await Promise.all([
    lineNumberingSourceZip.file("word/document.xml").async("text"),
    lineNumberingOutputZip.file("word/document.xml").async("text"),
  ]);
  const lineNumberingTags = (xml) => [...xml.matchAll(/<w:lnNumType\b[^>]*\/>/g)].map((match) => match[0]);
  const lineNumberingAttributes = (tag) => Object.fromEntries(
    [...tag.matchAll(/w:([\w-]+)="([^"]*)"/g)].map((match) => [match[1], match[1] === "restart" ? match[2] : Number(match[2])]),
  );
  const sourceLineNumberingTags = lineNumberingTags(lineNumberingSourceXml);
  const outputLineNumberingTags = lineNumberingTags(lineNumberingOutputXml);
  assert.equal(sourceLineNumberingTags.length, 2);
  assert.equal(outputLineNumberingTags.length, sourceLineNumberingTags.length);
  assert.deepEqual(lineNumberingAttributes(sourceLineNumberingTags[0]), { countBy: 5, start: 0, distance: 360, restart: "newPage" });
  assert.deepEqual(lineNumberingAttributes(outputLineNumberingTags[0]), { countBy: 10, start: 4, distance: 480, restart: "continuous" });
  assert.deepEqual(outputLineNumberingTags.slice(1), sourceLineNumberingTags.slice(1));
  const lineNumberingWorkflowOutputDocument = await DocumentFile.importDocx(await FileBlob.load(lineNumberingWorkflowOutputPath));
  assert.deepEqual(lineNumberingWorkflowOutputDocument.blocks[1]?.lineNumbering, replacementLineNumbering);
  assert.deepEqual(lineNumberingWorkflowOutputDocument.blocks[3]?.lineNumbering, { countBy: 2, start: 3, distance: 240, restart: "newSection" });
  const lineNumberingWorkflowRender = await verifyDocumentFile(lineNumberingWorkflowOutputPath, {
    outputDir: path.join(outputDir, "source-bound-section-line-numbering-render"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(lineNumberingWorkflowRender.summary.verifyOk, true);
  assert.equal(lineNumberingWorkflowRender.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const lineNumberingCliOutputPath = path.join(outputDir, "source-bound-section-line-numbering-cli.docx");
  const lineNumberingCliAuditPath = path.join(outputDir, "source-bound-section-line-numbering-cli-audit.json");
  const lineNumberingCliReplacement = { countBy: 3, distance: 240, restart: "newSection" };
  assert.deepEqual(parseSectionLineNumberingEditCli([
    lineNumberingSourcePath,
    lineNumberingCliOutputPath,
    lineNumberingCliAuditPath,
    "1",
    JSON.stringify(sourceLineNumbering),
    JSON.stringify(lineNumberingCliReplacement),
  ]), {
    inputPath: lineNumberingSourcePath,
    outputPath: lineNumberingCliOutputPath,
    auditPath: lineNumberingCliAuditPath,
    sectionBlockIndex: 1,
    expectedLineNumbering: sourceLineNumbering,
    replacementLineNumbering: lineNumberingCliReplacement,
  });
  const lineNumberingCliProcess = spawnSync(process.execPath, [
    path.join(repoRoot, "skills", "documents", "skills", "documents", "examples", "officekit-section-line-numbering-edit-workflow.mjs"),
    lineNumberingSourcePath,
    lineNumberingCliOutputPath,
    lineNumberingCliAuditPath,
    "1",
    JSON.stringify(sourceLineNumbering),
    JSON.stringify(lineNumberingCliReplacement),
  ], { cwd: repoRoot, encoding: "utf8" });
  assert.equal(lineNumberingCliProcess.status, 0, lineNumberingCliProcess.stderr);
  assert.deepEqual(JSON.parse(lineNumberingCliProcess.stdout), {
    outputPath: lineNumberingCliOutputPath,
    auditPath: lineNumberingCliAuditPath,
    outputSha256: createHash("sha256").update(await fs.readFile(lineNumberingCliOutputPath)).digest("hex"),
    changedParts: ["word/document.xml"],
  });
  const lineNumberingCliDocument = await DocumentFile.importDocx(await FileBlob.load(lineNumberingCliOutputPath));
  assert.deepEqual(lineNumberingCliDocument.blocks[1]?.lineNumbering, lineNumberingCliReplacement);

  await assert.rejects(
    () => editImportedSectionLineNumbering({
      inputPath: lineNumberingSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-line-numbering-mismatched.docx"),
      auditPath: path.join(outputDir, "source-bound-section-line-numbering-mismatched-audit.json"),
      sectionBlockIndex: 1,
      expectedLineNumbering: { ...sourceLineNumbering, countBy: 1 },
      replacementLineNumbering,
    }),
    /lineNumbering does not match the expected source value/,
  );
  await assert.rejects(
    () => editImportedSectionLineNumbering({
      inputPath: lineNumberingSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-line-numbering-invalid.docx"),
      auditPath: path.join(outputDir, "source-bound-section-line-numbering-invalid-audit.json"),
      sectionBlockIndex: 1,
      expectedLineNumbering: sourceLineNumbering,
      replacementLineNumbering: { start: 1 },
    }),
    /replacementLineNumbering\.countBy is required/,
  );
  await assert.rejects(
    () => editImportedSectionLineNumbering({
      inputPath: lineNumberingSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-line-numbering-noop.docx"),
      auditPath: path.join(outputDir, "source-bound-section-line-numbering-noop-audit.json"),
      sectionBlockIndex: 1,
      expectedLineNumbering: sourceLineNumbering,
      replacementLineNumbering: sourceLineNumbering,
    }),
    /replacementLineNumbering must differ from expectedLineNumbering/,
  );
  const implicitLineNumberingSourcePath = path.join(outputDir, "source-bound-section-line-numbering-implicit-source.docx");
  const implicitLineNumberingZip = await JSZip.loadAsync(lineNumberingSourceBytes);
  const implicitLineNumberingXml = await implicitLineNumberingZip.file("word/document.xml").async("text");
  implicitLineNumberingZip.file("word/document.xml", implicitLineNumberingXml.replace(sourceLineNumberingTags[0], "<w:lnNumType/>"));
  await fs.writeFile(implicitLineNumberingSourcePath, await implicitLineNumberingZip.generateAsync({ type: "nodebuffer" }), { flag: "wx" });
  const implicitLineNumberingOutputPath = path.join(outputDir, "source-bound-section-line-numbering-implicit-edited.docx");
  const implicitLineNumberingAuditPath = path.join(outputDir, "source-bound-section-line-numbering-implicit-edited-audit.json");
  const implicitLineNumbering = await editImportedSectionLineNumbering({
    inputPath: implicitLineNumberingSourcePath,
    outputPath: implicitLineNumberingOutputPath,
    auditPath: implicitLineNumberingAuditPath,
    sectionBlockIndex: 1,
    expectedLineNumbering: { countBy: 1 },
    replacementLineNumbering: { countBy: 4, restart: "continuous" },
  });
  assert.deepEqual(implicitLineNumbering.audit.validation.reimport.lineNumbering, { countBy: 4, restart: "continuous" });
  const nonCanonicalLineNumberingSourcePath = path.join(outputDir, "source-bound-section-line-numbering-noncanonical-source.docx");
  const nonCanonicalLineNumberingZip = await JSZip.loadAsync(lineNumberingSourceBytes);
  const nonCanonicalLineNumberingXml = await nonCanonicalLineNumberingZip.file("word/document.xml").async("text");
  nonCanonicalLineNumberingZip.file("word/document.xml", nonCanonicalLineNumberingXml.replace(/w:countBy="5"/, 'w:countBy="05"'));
  await fs.writeFile(nonCanonicalLineNumberingSourcePath, await nonCanonicalLineNumberingZip.generateAsync({ type: "nodebuffer" }), { flag: "wx" });
  await assert.rejects(
    () => editImportedSectionLineNumbering({
      inputPath: nonCanonicalLineNumberingSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-line-numbering-noncanonical.docx"),
      auditPath: path.join(outputDir, "source-bound-section-line-numbering-noncanonical-audit.json"),
      sectionBlockIndex: 1,
      expectedLineNumbering: sourceLineNumbering,
      replacementLineNumbering,
    }),
    /canonical unsigned integer/,
  );
  assert.deepEqual(await fs.readFile(lineNumberingSourcePath), lineNumberingSourceBytes);

  const {
    editImportedSectionColumns,
    parseSectionColumnsEditCli,
    sectionColumnsCliOutput,
  } = await import(
    "../skills/documents/skills/documents/examples/officekit-section-columns-edit-workflow.mjs"
  );
  const columnsSourceDocument = DocumentModel.create({ name: "Source-bound section columns edit", blocks: [] });
  columnsSourceDocument.addParagraph("Prelude for the canonical section columns transaction.");
  columnsSourceDocument.addSection({
    breakType: "nextPage",
    columns: { count: 2, spacing: 720, separator: true },
  });
  columnsSourceDocument.addParagraph("Only this section's equal-width column profile may change.");
  columnsSourceDocument.addSection({
    breakType: "nextPage",
    columns: { definitions: [{ width: 3000, spacing: 720 }, { width: 5640, spacing: 0 }], separator: true },
  });
  columnsSourceDocument.addParagraph("The custom-width sibling section is a raw-XML canary.");
  const columnsSourcePath = path.join(outputDir, "source-bound-section-columns-source.docx");
  await (await DocumentFile.exportDocx(columnsSourceDocument)).save(columnsSourcePath);
  const columnsSourceBytes = await fs.readFile(columnsSourcePath);
  const columnsImported = await DocumentFile.importDocx(await FileBlob.load(columnsSourcePath));
  const sourceColumns = { count: 2, spacing: 720, separator: true };
  const replacementColumns = { count: 3, spacing: 360, separator: false };
  const siblingCustomColumns = { definitions: [{ width: 3000, spacing: 720 }, { width: 5640, spacing: 0 }], separator: true };
  assert.deepEqual(columnsImported.blocks[1]?.columns, sourceColumns);
  assert.deepEqual(columnsImported.blocks[3]?.columns, siblingCustomColumns);

  const columnsWorkflowOutputPath = path.join(outputDir, "source-bound-section-columns-edited.docx");
  const columnsWorkflowAuditPath = path.join(outputDir, "source-bound-section-columns-edited-audit.json");
  const columnsWorkflow = await editImportedSectionColumns({
    inputPath: columnsSourcePath,
    outputPath: columnsWorkflowOutputPath,
    auditPath: columnsWorkflowAuditPath,
    sectionBlockIndex: 1,
    expectedColumns: sourceColumns,
    replacementColumns,
  });
  assert.equal(columnsWorkflow.audit.provider.actual, "office-kit");
  assert.equal(columnsWorkflow.audit.provider.silentFallback, false);
  assert.deepEqual(columnsWorkflow.audit.savePolicy, { strategy: "rewrite", noReplace: true });
  assert.equal(columnsWorkflow.audit.operation.type, "source-bound-section-columns-edit");
  assert.deepEqual(columnsWorkflow.audit.operation.target, {
    id: columnsImported.blocks[1].id,
    blockIndex: 1,
    sectionOrdinal: 0,
  });
  assert.deepEqual(columnsWorkflow.audit.operation.sourceColumns, sourceColumns);
  assert.deepEqual(columnsWorkflow.audit.operation.replacementColumns, replacementColumns);
  assert.deepEqual(columnsWorkflow.audit.validation.changedParts, ["word/document.xml"]);
  assert.equal(columnsWorkflow.audit.validation.columnsXmlResidual.ok, true);
  assert.deepEqual(columnsWorkflow.audit.validation.reimport.columns, replacementColumns);
  assert.equal(columnsWorkflow.audit.validation.reimport.editable, true);
  assert.equal(columnsWorkflow.audit.validation.nativeRenderRequired, true);
  assert.deepEqual(sectionColumnsCliOutput(columnsWorkflow).changedParts, ["word/document.xml"]);
  assert.deepEqual(await fs.readFile(columnsSourcePath), columnsSourceBytes);

  const columnsWorkflowOutputBytes = await fs.readFile(columnsWorkflowOutputPath);
  const [columnsSourceZip, columnsOutputZip] = await Promise.all([
    JSZip.loadAsync(columnsSourceBytes),
    JSZip.loadAsync(columnsWorkflowOutputBytes),
  ]);
  const columnsParts = Object.keys(columnsSourceZip.files).filter((partPath) => !columnsSourceZip.files[partPath].dir).sort();
  assert.deepEqual(
    Object.keys(columnsOutputZip.files).filter((partPath) => !columnsOutputZip.files[partPath].dir).sort(),
    columnsParts,
  );
  for (const partPath of columnsParts) {
    if (partPath === "word/document.xml") continue;
    assert.deepEqual(
      Buffer.from(await columnsOutputZip.file(partPath).async("uint8array")),
      Buffer.from(await columnsSourceZip.file(partPath).async("uint8array")),
      `Only word/document.xml may change; ${partPath} drifted.`,
    );
  }
  const [columnsSourceXml, columnsOutputXml] = await Promise.all([
    columnsSourceZip.file("word/document.xml").async("text"),
    columnsOutputZip.file("word/document.xml").async("text"),
  ]);
  const columnsMarkup = (xml) => [...xml.matchAll(/<w:cols\b[^>]*?\/>|<w:cols\b[^>]*>[\s\S]*?<\/w:cols>/g)].map((match) => match[0]);
  const sourceColumnsMarkup = columnsMarkup(columnsSourceXml);
  const outputColumnsMarkup = columnsMarkup(columnsOutputXml);
  assert.equal(sourceColumnsMarkup.length, 2);
  assert.equal(outputColumnsMarkup.length, sourceColumnsMarkup.length);
  assert.match(sourceColumnsMarkup[0], /w:equalWidth="(?:true|1)"/);
  assert.match(sourceColumnsMarkup[0], /w:num="2"/);
  assert.match(outputColumnsMarkup[0], /w:equalWidth="(?:true|1)"/);
  assert.match(outputColumnsMarkup[0], /w:num="3"/);
  assert.match(outputColumnsMarkup[0], /w:space="360"/);
  assert.match(outputColumnsMarkup[0], /w:sep="(?:false|0)"/);
  assert.equal(outputColumnsMarkup[1], sourceColumnsMarkup[1]);
  const columnsWorkflowOutputDocument = await DocumentFile.importDocx(await FileBlob.load(columnsWorkflowOutputPath));
  assert.deepEqual(columnsWorkflowOutputDocument.blocks[1]?.columns, replacementColumns);
  assert.deepEqual(columnsWorkflowOutputDocument.blocks[3]?.columns, siblingCustomColumns);
  const columnsWorkflowRender = await verifyDocumentFile(columnsWorkflowOutputPath, {
    outputDir: path.join(outputDir, "source-bound-section-columns-render"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(columnsWorkflowRender.summary.verifyOk, true);
  assert.equal(columnsWorkflowRender.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const columnsCliOutputPath = path.join(outputDir, "source-bound-section-columns-cli.docx");
  const columnsCliAuditPath = path.join(outputDir, "source-bound-section-columns-cli-audit.json");
  const columnsCliReplacement = { count: 2, spacing: 480, separator: false };
  assert.deepEqual(parseSectionColumnsEditCli([
    columnsSourcePath,
    columnsCliOutputPath,
    columnsCliAuditPath,
    "1",
    JSON.stringify(sourceColumns),
    JSON.stringify(columnsCliReplacement),
  ]), {
    inputPath: columnsSourcePath,
    outputPath: columnsCliOutputPath,
    auditPath: columnsCliAuditPath,
    sectionBlockIndex: 1,
    expectedColumns: sourceColumns,
    replacementColumns: columnsCliReplacement,
  });
  const columnsCliProcess = spawnSync(process.execPath, [
    path.join(repoRoot, "skills", "documents", "skills", "documents", "examples", "officekit-section-columns-edit-workflow.mjs"),
    columnsSourcePath,
    columnsCliOutputPath,
    columnsCliAuditPath,
    "1",
    JSON.stringify(sourceColumns),
    JSON.stringify(columnsCliReplacement),
  ], { cwd: repoRoot, encoding: "utf8" });
  assert.equal(columnsCliProcess.status, 0, columnsCliProcess.stderr);
  assert.deepEqual(JSON.parse(columnsCliProcess.stdout), {
    outputPath: columnsCliOutputPath,
    auditPath: columnsCliAuditPath,
    outputSha256: createHash("sha256").update(await fs.readFile(columnsCliOutputPath)).digest("hex"),
    changedParts: ["word/document.xml"],
  });
  assert.deepEqual((await DocumentFile.importDocx(await FileBlob.load(columnsCliOutputPath))).blocks[1]?.columns, columnsCliReplacement);

  const customColumnsSourceDocument = DocumentModel.create({ name: "Source-bound custom-width section columns edit", blocks: [] });
  customColumnsSourceDocument.addParagraph("Prelude for custom-width section columns.");
  customColumnsSourceDocument.addSection({
    breakType: "nextPage",
    columns: siblingCustomColumns,
  });
  customColumnsSourceDocument.addParagraph("Only this section's explicit column definitions may change.");
  const customColumnsSourcePath = path.join(outputDir, "source-bound-section-columns-custom-source.docx");
  await (await DocumentFile.exportDocx(customColumnsSourceDocument)).save(customColumnsSourcePath);
  const customColumnsOutputPath = path.join(outputDir, "source-bound-section-columns-custom-edited.docx");
  const customColumnsAuditPath = path.join(outputDir, "source-bound-section-columns-custom-edited-audit.json");
  const replacementCustomColumns = { definitions: [{ width: 3200, spacing: 360 }, { width: 5680, spacing: 0 }], separator: false };
  const customColumnsWorkflow = await editImportedSectionColumns({
    inputPath: customColumnsSourcePath,
    outputPath: customColumnsOutputPath,
    auditPath: customColumnsAuditPath,
    sectionBlockIndex: 1,
    expectedColumns: siblingCustomColumns,
    replacementColumns: replacementCustomColumns,
  });
  assert.deepEqual(customColumnsWorkflow.audit.validation.changedParts, ["word/document.xml"]);
  assert.deepEqual((await DocumentFile.importDocx(await FileBlob.load(customColumnsOutputPath))).blocks[1]?.columns, replacementCustomColumns);

  await assert.rejects(
    () => editImportedSectionColumns({
      inputPath: columnsSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-columns-mismatched.docx"),
      auditPath: path.join(outputDir, "source-bound-section-columns-mismatched-audit.json"),
      sectionBlockIndex: 1,
      expectedColumns: { ...sourceColumns, count: 4 },
      replacementColumns,
    }),
    /columns do not match the expected source value/,
  );
  await assert.rejects(
    () => editImportedSectionColumns({
      inputPath: columnsSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-columns-invalid.docx"),
      auditPath: path.join(outputDir, "source-bound-section-columns-invalid-audit.json"),
      sectionBlockIndex: 1,
      expectedColumns: sourceColumns,
      replacementColumns: { count: 3, spacing: 360 },
    }),
    /replacementColumns\.separator is required/,
  );
  await assert.rejects(
    () => editImportedSectionColumns({
      inputPath: columnsSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-columns-shape-change.docx"),
      auditPath: path.join(outputDir, "source-bound-section-columns-shape-change-audit.json"),
      sectionBlockIndex: 1,
      expectedColumns: sourceColumns,
      replacementColumns: { definitions: siblingCustomColumns.definitions, separator: false },
    }),
    /must retain the source equal-width or custom-width profile/,
  );
  await assert.rejects(
    () => editImportedSectionColumns({
      inputPath: columnsSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-columns-noop.docx"),
      auditPath: path.join(outputDir, "source-bound-section-columns-noop-audit.json"),
      sectionBlockIndex: 1,
      expectedColumns: sourceColumns,
      replacementColumns: sourceColumns,
    }),
    /replacementColumns must differ from expectedColumns/,
  );
  const nonCanonicalColumnsSourcePath = path.join(outputDir, "source-bound-section-columns-noncanonical-source.docx");
  const nonCanonicalColumnsZip = await JSZip.loadAsync(columnsSourceBytes);
  nonCanonicalColumnsZip.file("word/document.xml", columnsSourceXml.replace(/w:num="2"/, 'w:num="02"'));
  await fs.writeFile(nonCanonicalColumnsSourcePath, await nonCanonicalColumnsZip.generateAsync({ type: "nodebuffer" }), { flag: "wx" });
  await assert.rejects(
    () => editImportedSectionColumns({
      inputPath: nonCanonicalColumnsSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-columns-noncanonical.docx"),
      auditPath: path.join(outputDir, "source-bound-section-columns-noncanonical-audit.json"),
      sectionBlockIndex: 1,
      expectedColumns: sourceColumns,
      replacementColumns,
    }),
    /canonical unsigned integer/,
  );
  assert.deepEqual(await fs.readFile(columnsSourcePath), columnsSourceBytes);

  const {
    editImportedSectionBreakType,
    parseSectionBreakTypeEditCli,
    sectionBreakTypeCliOutput,
  } = await import(
    "../skills/documents/skills/documents/examples/officekit-section-break-edit-workflow.mjs"
  );
  const breakTypeSourceDocument = DocumentModel.create({ name: "Source-bound section break-type edit", blocks: [] });
  breakTypeSourceDocument.addParagraph("Prelude for the canonical section-break transaction.");
  breakTypeSourceDocument.addSection({ breakType: "nextPage" });
  breakTypeSourceDocument.addParagraph("Only this section's native break type may change.");
  breakTypeSourceDocument.addSection({ breakType: "evenPage" });
  breakTypeSourceDocument.addParagraph("The sibling section is a raw-XML canary.");
  const breakTypeSourcePath = path.join(outputDir, "source-bound-section-break-type-source.docx");
  await (await DocumentFile.exportDocx(breakTypeSourceDocument)).save(breakTypeSourcePath);
  const breakTypeSourceBytes = await fs.readFile(breakTypeSourcePath);
  const breakTypeImported = await DocumentFile.importDocx(await FileBlob.load(breakTypeSourcePath));
  assert.equal(breakTypeImported.blocks[1]?.breakType, "nextPage");
  assert.equal(breakTypeImported.blocks[3]?.breakType, "evenPage");

  const breakTypeWorkflowOutputPath = path.join(outputDir, "source-bound-section-break-type-edited.docx");
  const breakTypeWorkflowAuditPath = path.join(outputDir, "source-bound-section-break-type-edited-audit.json");
  const breakTypeWorkflow = await editImportedSectionBreakType({
    inputPath: breakTypeSourcePath,
    outputPath: breakTypeWorkflowOutputPath,
    auditPath: breakTypeWorkflowAuditPath,
    sectionBlockIndex: 1,
    expectedBreakType: "nextPage",
    replacementBreakType: "continuous",
  });
  assert.equal(breakTypeWorkflow.audit.provider.actual, "office-kit");
  assert.equal(breakTypeWorkflow.audit.provider.silentFallback, false);
  assert.deepEqual(breakTypeWorkflow.audit.savePolicy, { strategy: "rewrite", noReplace: true });
  assert.equal(breakTypeWorkflow.audit.operation.type, "source-bound-section-break-type-edit");
  assert.deepEqual(breakTypeWorkflow.audit.operation.target, {
    id: breakTypeImported.blocks[1].id,
    blockIndex: 1,
    sectionOrdinal: 0,
  });
  assert.equal(breakTypeWorkflow.audit.operation.sourceBreakType, "nextPage");
  assert.equal(breakTypeWorkflow.audit.operation.replacementBreakType, "continuous");
  assert.deepEqual(breakTypeWorkflow.audit.validation.changedParts, ["word/document.xml"]);
  assert.equal(breakTypeWorkflow.audit.validation.sectionTypeXmlResidual.ok, true);
  assert.equal(breakTypeWorkflow.audit.validation.reimport.breakType, "continuous");
  assert.equal(breakTypeWorkflow.audit.validation.reimport.editable, true);
  assert.equal(breakTypeWorkflow.audit.validation.nativeRenderRequired, true);
  assert.deepEqual(sectionBreakTypeCliOutput(breakTypeWorkflow).changedParts, ["word/document.xml"]);
  assert.deepEqual(await fs.readFile(breakTypeSourcePath), breakTypeSourceBytes);

  const breakTypeWorkflowOutputBytes = await fs.readFile(breakTypeWorkflowOutputPath);
  const [breakTypeSourceZip, breakTypeOutputZip] = await Promise.all([
    JSZip.loadAsync(breakTypeSourceBytes),
    JSZip.loadAsync(breakTypeWorkflowOutputBytes),
  ]);
  const breakTypeParts = Object.keys(breakTypeSourceZip.files).filter((partPath) => !breakTypeSourceZip.files[partPath].dir).sort();
  assert.deepEqual(
    Object.keys(breakTypeOutputZip.files).filter((partPath) => !breakTypeOutputZip.files[partPath].dir).sort(),
    breakTypeParts,
  );
  for (const partPath of breakTypeParts) {
    if (partPath === "word/document.xml") continue;
    assert.deepEqual(
      Buffer.from(await breakTypeOutputZip.file(partPath).async("uint8array")),
      Buffer.from(await breakTypeSourceZip.file(partPath).async("uint8array")),
      `Only word/document.xml may change; ${partPath} drifted.`,
    );
  }
  const [breakTypeSourceXml, breakTypeOutputXml] = await Promise.all([
    breakTypeSourceZip.file("word/document.xml").async("text"),
    breakTypeOutputZip.file("word/document.xml").async("text"),
  ]);
  const sectionTypeMarkup = (xml) => [...xml.matchAll(/<w:type\b[^>]*\/>/g)].map((match) => match[0]);
  const sourceTypeMarkup = sectionTypeMarkup(breakTypeSourceXml);
  const outputTypeMarkup = sectionTypeMarkup(breakTypeOutputXml);
  assert.equal(sourceTypeMarkup.length, 3);
  assert.equal(outputTypeMarkup.length, sourceTypeMarkup.length);
  assert.match(sourceTypeMarkup[0], /w:val="nextPage"/);
  assert.match(outputTypeMarkup[0], /w:val="continuous"/);
  assert.equal(outputTypeMarkup[1], sourceTypeMarkup[1]);
  assert.equal(outputTypeMarkup[2], sourceTypeMarkup[2]);
  const breakTypeWorkflowOutputDocument = await DocumentFile.importDocx(await FileBlob.load(breakTypeWorkflowOutputPath));
  assert.equal(breakTypeWorkflowOutputDocument.blocks[1]?.breakType, "continuous");
  assert.equal(breakTypeWorkflowOutputDocument.blocks[3]?.breakType, "evenPage");
  const breakTypeWorkflowRender = await verifyDocumentFile(breakTypeWorkflowOutputPath, {
    outputDir: path.join(outputDir, "source-bound-section-break-type-render"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(breakTypeWorkflowRender.summary.verifyOk, true);
  assert.equal(breakTypeWorkflowRender.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const breakTypeCliOutputPath = path.join(outputDir, "source-bound-section-break-type-cli.docx");
  const breakTypeCliAuditPath = path.join(outputDir, "source-bound-section-break-type-cli-audit.json");
  assert.deepEqual(parseSectionBreakTypeEditCli([
    breakTypeSourcePath,
    breakTypeCliOutputPath,
    breakTypeCliAuditPath,
    "1",
    "nextPage",
    "oddPage",
  ]), {
    inputPath: breakTypeSourcePath,
    outputPath: breakTypeCliOutputPath,
    auditPath: breakTypeCliAuditPath,
    sectionBlockIndex: 1,
    expectedBreakType: "nextPage",
    replacementBreakType: "oddPage",
  });
  const breakTypeCliProcess = spawnSync(process.execPath, [
    path.join(repoRoot, "skills", "documents", "skills", "documents", "examples", "officekit-section-break-edit-workflow.mjs"),
    breakTypeSourcePath,
    breakTypeCliOutputPath,
    breakTypeCliAuditPath,
    "1",
    "nextPage",
    "oddPage",
  ], { cwd: repoRoot, encoding: "utf8" });
  assert.equal(breakTypeCliProcess.status, 0, breakTypeCliProcess.stderr);
  assert.deepEqual(JSON.parse(breakTypeCliProcess.stdout), {
    outputPath: breakTypeCliOutputPath,
    auditPath: breakTypeCliAuditPath,
    outputSha256: createHash("sha256").update(await fs.readFile(breakTypeCliOutputPath)).digest("hex"),
    changedParts: ["word/document.xml"],
  });
  assert.equal((await DocumentFile.importDocx(await FileBlob.load(breakTypeCliOutputPath))).blocks[1]?.breakType, "oddPage");

  await assert.rejects(
    () => editImportedSectionBreakType({
      inputPath: breakTypeSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-break-type-mismatched.docx"),
      auditPath: path.join(outputDir, "source-bound-section-break-type-mismatched-audit.json"),
      sectionBlockIndex: 1,
      expectedBreakType: "evenPage",
      replacementBreakType: "continuous",
    }),
    /breakType does not match the expected source value/,
  );
  await assert.rejects(
    () => editImportedSectionBreakType({
      inputPath: breakTypeSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-break-type-invalid.docx"),
      auditPath: path.join(outputDir, "source-bound-section-break-type-invalid-audit.json"),
      sectionBlockIndex: 1,
      expectedBreakType: "nextPage",
      replacementBreakType: "newPage",
    }),
    /replacementBreakType must be nextPage, continuous, evenPage, or oddPage/,
  );
  await assert.rejects(
    () => editImportedSectionBreakType({
      inputPath: breakTypeSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-break-type-noop.docx"),
      auditPath: path.join(outputDir, "source-bound-section-break-type-noop-audit.json"),
      sectionBlockIndex: 1,
      expectedBreakType: "nextPage",
      replacementBreakType: "nextPage",
    }),
    /replacementBreakType must differ from expectedBreakType/,
  );
  const missingSectionTypeSourcePath = path.join(outputDir, "source-bound-section-break-type-missing-source.docx");
  const missingSectionTypeZip = await JSZip.loadAsync(breakTypeSourceBytes);
  missingSectionTypeZip.file("word/document.xml", breakTypeSourceXml.replace(sourceTypeMarkup[0], ""));
  await fs.writeFile(missingSectionTypeSourcePath, await missingSectionTypeZip.generateAsync({ type: "nodebuffer" }), { flag: "wx" });
  await assert.rejects(
    () => editImportedSectionBreakType({
      inputPath: missingSectionTypeSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-break-type-missing.docx"),
      auditPath: path.join(outputDir, "source-bound-section-break-type-missing-audit.json"),
      sectionBlockIndex: 1,
      expectedBreakType: "nextPage",
      replacementBreakType: "continuous",
    }),
    /exactly one canonical w:type leaf/,
  );
  assert.deepEqual(await fs.readFile(breakTypeSourcePath), breakTypeSourceBytes);

  const {
    editImportedTableColumnWidths,
    parseTableColumnWidthsEditCli,
    tableColumnWidthsCliOutput,
  } = await import(
    "../skills/documents/skills/documents/examples/officekit-table-column-widths-edit-workflow.mjs"
  );
  const tableWidthSourceDocument = DocumentModel.create({ name: "Source-bound table column-width edit", blocks: [] });
  tableWidthSourceDocument.addParagraph("The first table is the target fixed-layout grid.");
  tableWidthSourceDocument.addTable({
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
  });
  tableWidthSourceDocument.addParagraph("The second table is a raw-XML canary.");
  tableWidthSourceDocument.addTable({
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
  const tableWidthSourcePath = path.join(outputDir, "source-bound-table-column-widths-source.docx");
  await (await DocumentFile.exportDocx(tableWidthSourceDocument)).save(tableWidthSourcePath);
  const tableWidthSourceBytes = await fs.readFile(tableWidthSourcePath);
  const tableWidthImported = await DocumentFile.importDocx(await FileBlob.load(tableWidthSourcePath));
  const sourceColumnWidths = [2100, 4500, 2700];
  const replacementColumnWidths = [3000, 3600, 2700];
  const siblingColumnWidths = [3300, 6000];
  assert.deepEqual(tableWidthImported.blocks[1]?.columnWidthsDxa, sourceColumnWidths);
  assert.deepEqual(tableWidthImported.blocks[3]?.columnWidthsDxa, siblingColumnWidths);
  assert.equal(tableWidthImported.blocks[1]?.sourceBound, true);

  const tableWidthWorkflowOutputPath = path.join(outputDir, "source-bound-table-column-widths-edited.docx");
  const tableWidthWorkflowAuditPath = path.join(outputDir, "source-bound-table-column-widths-edited-audit.json");
  const tableWidthWorkflow = await editImportedTableColumnWidths({
    inputPath: tableWidthSourcePath,
    outputPath: tableWidthWorkflowOutputPath,
    auditPath: tableWidthWorkflowAuditPath,
    tableBlockIndex: 1,
    expectedColumnWidthsDxa: sourceColumnWidths,
    replacementColumnWidthsDxa: replacementColumnWidths,
  });
  assert.equal(tableWidthWorkflow.audit.provider.actual, "office-kit");
  assert.equal(tableWidthWorkflow.audit.provider.silentFallback, false);
  assert.deepEqual(tableWidthWorkflow.audit.savePolicy, { strategy: "rewrite", noReplace: true });
  assert.equal(tableWidthWorkflow.audit.operation.type, "source-bound-table-column-widths-edit");
  assert.deepEqual(tableWidthWorkflow.audit.operation.target, {
    id: tableWidthImported.blocks[1].id,
    blockIndex: 1,
    tableOrdinal: 0,
  });
  assert.equal(tableWidthWorkflow.audit.operation.tableWidthDxa, 9300);
  assert.deepEqual(tableWidthWorkflow.audit.operation.sourceColumnWidthsDxa, sourceColumnWidths);
  assert.deepEqual(tableWidthWorkflow.audit.operation.replacementColumnWidthsDxa, replacementColumnWidths);
  assert.deepEqual(tableWidthWorkflow.audit.validation.changedParts, ["word/document.xml"]);
  assert.equal(tableWidthWorkflow.audit.validation.tableWidthXmlResidual.ok, true);
  assert.deepEqual(tableWidthWorkflow.audit.validation.reimport.columnWidthsDxa, replacementColumnWidths);
  assert.equal(tableWidthWorkflow.audit.validation.nativeRenderRequired, true);
  assert.deepEqual(tableColumnWidthsCliOutput(tableWidthWorkflow).changedParts, ["word/document.xml"]);
  assert.deepEqual(await fs.readFile(tableWidthSourcePath), tableWidthSourceBytes);

  const tableWidthWorkflowOutputBytes = await fs.readFile(tableWidthWorkflowOutputPath);
  const [tableWidthSourceZip, tableWidthOutputZip] = await Promise.all([
    JSZip.loadAsync(tableWidthSourceBytes),
    JSZip.loadAsync(tableWidthWorkflowOutputBytes),
  ]);
  const tableWidthParts = Object.keys(tableWidthSourceZip.files).filter((partPath) => !tableWidthSourceZip.files[partPath].dir).sort();
  assert.deepEqual(
    Object.keys(tableWidthOutputZip.files).filter((partPath) => !tableWidthOutputZip.files[partPath].dir).sort(),
    tableWidthParts,
  );
  for (const partPath of tableWidthParts) {
    if (partPath === "word/document.xml") continue;
    assert.deepEqual(
      Buffer.from(await tableWidthOutputZip.file(partPath).async("uint8array")),
      Buffer.from(await tableWidthSourceZip.file(partPath).async("uint8array")),
      `Only word/document.xml may change; ${partPath} drifted.`,
    );
  }
  const [tableWidthSourceXml, tableWidthOutputXml] = await Promise.all([
    tableWidthSourceZip.file("word/document.xml").async("text"),
    tableWidthOutputZip.file("word/document.xml").async("text"),
  ]);
  const flatTableMarkup = (xml) => [...xml.matchAll(/<w:tbl\b[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
  const tableGridWidths = (tableXml) => [...tableXml.matchAll(/<w:gridCol\b[^>]*w:w="(\d+)"[^>]*\/>/g)].map((match) => Number(match[1]));
  const tableCellWidths = (tableXml) => [...tableXml.matchAll(/<w:tcW\b[^>]*w:w="(\d+)"[^>]*\/>/g)].map((match) => Number(match[1]));
  const maskTableWidths = (tableXml) => tableXml
    .replace(/<w:gridCol\b[^>]*\/>/g, '<w:gridCol w:w="officeKitWidthMasked"/>')
    .replace(/<w:tcW\b[^>]*\/>/g, '<w:tcW w:type="dxa" w:w="officeKitWidthMasked"/>');
  const sourceTables = flatTableMarkup(tableWidthSourceXml);
  const outputTables = flatTableMarkup(tableWidthOutputXml);
  assert.equal(sourceTables.length, 2);
  assert.equal(outputTables.length, sourceTables.length);
  assert.deepEqual(tableGridWidths(sourceTables[0]), sourceColumnWidths);
  assert.deepEqual(tableGridWidths(outputTables[0]), replacementColumnWidths);
  assert.deepEqual(tableCellWidths(sourceTables[0]), [...sourceColumnWidths, ...sourceColumnWidths, ...sourceColumnWidths]);
  assert.deepEqual(tableCellWidths(outputTables[0]), [...replacementColumnWidths, ...replacementColumnWidths, ...replacementColumnWidths]);
  assert.equal(maskTableWidths(outputTables[0]), maskTableWidths(sourceTables[0]));
  assert.equal(outputTables[1], sourceTables[1]);
  const tableWidthWorkflowOutputDocument = await DocumentFile.importDocx(await FileBlob.load(tableWidthWorkflowOutputPath));
  assert.deepEqual(tableWidthWorkflowOutputDocument.blocks[1]?.columnWidthsDxa, replacementColumnWidths);
  assert.deepEqual(tableWidthWorkflowOutputDocument.blocks[3]?.columnWidthsDxa, siblingColumnWidths);
  const tableWidthWorkflowRender = await verifyDocumentFile(tableWidthWorkflowOutputPath, {
    outputDir: path.join(outputDir, "source-bound-table-column-widths-render"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(tableWidthWorkflowRender.summary.verifyOk, true);
  assert.equal(tableWidthWorkflowRender.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const tableWidthCliOutputPath = path.join(outputDir, "source-bound-table-column-widths-cli.docx");
  const tableWidthCliAuditPath = path.join(outputDir, "source-bound-table-column-widths-cli-audit.json");
  const tableWidthCliReplacement = [2600, 4000, 2700];
  assert.deepEqual(parseTableColumnWidthsEditCli([
    tableWidthSourcePath,
    tableWidthCliOutputPath,
    tableWidthCliAuditPath,
    "1",
    JSON.stringify(sourceColumnWidths),
    JSON.stringify(tableWidthCliReplacement),
  ]), {
    inputPath: tableWidthSourcePath,
    outputPath: tableWidthCliOutputPath,
    auditPath: tableWidthCliAuditPath,
    tableBlockIndex: 1,
    expectedColumnWidthsDxa: sourceColumnWidths,
    replacementColumnWidthsDxa: tableWidthCliReplacement,
  });
  const tableWidthCliProcess = spawnSync(process.execPath, [
    path.join(repoRoot, "skills", "documents", "skills", "documents", "examples", "officekit-table-column-widths-edit-workflow.mjs"),
    tableWidthSourcePath,
    tableWidthCliOutputPath,
    tableWidthCliAuditPath,
    "1",
    JSON.stringify(sourceColumnWidths),
    JSON.stringify(tableWidthCliReplacement),
  ], { cwd: repoRoot, encoding: "utf8" });
  assert.equal(tableWidthCliProcess.status, 0, tableWidthCliProcess.stderr);
  assert.deepEqual(JSON.parse(tableWidthCliProcess.stdout), {
    outputPath: tableWidthCliOutputPath,
    auditPath: tableWidthCliAuditPath,
    outputSha256: createHash("sha256").update(await fs.readFile(tableWidthCliOutputPath)).digest("hex"),
    changedParts: ["word/document.xml"],
  });
  assert.deepEqual((await DocumentFile.importDocx(await FileBlob.load(tableWidthCliOutputPath))).blocks[1]?.columnWidthsDxa, tableWidthCliReplacement);

  await assert.rejects(
    () => editImportedTableColumnWidths({
      inputPath: tableWidthSourcePath,
      outputPath: path.join(outputDir, "source-bound-table-column-widths-mismatched.docx"),
      auditPath: path.join(outputDir, "source-bound-table-column-widths-mismatched-audit.json"),
      tableBlockIndex: 1,
      expectedColumnWidthsDxa: [2100, 4400, 2800],
      replacementColumnWidthsDxa: replacementColumnWidths,
    }),
    /column widths do not match the expected source value/,
  );
  await assert.rejects(
    () => editImportedTableColumnWidths({
      inputPath: tableWidthSourcePath,
      outputPath: path.join(outputDir, "source-bound-table-column-widths-noop.docx"),
      auditPath: path.join(outputDir, "source-bound-table-column-widths-noop-audit.json"),
      tableBlockIndex: 1,
      expectedColumnWidthsDxa: sourceColumnWidths,
      replacementColumnWidthsDxa: sourceColumnWidths,
    }),
    /replacementColumnWidthsDxa must differ from expectedColumnWidthsDxa/,
  );
  await assert.rejects(
    () => editImportedTableColumnWidths({
      inputPath: tableWidthSourcePath,
      outputPath: path.join(outputDir, "source-bound-table-column-widths-total.docx"),
      auditPath: path.join(outputDir, "source-bound-table-column-widths-total-audit.json"),
      tableBlockIndex: 1,
      expectedColumnWidthsDxa: sourceColumnWidths,
      replacementColumnWidthsDxa: [3000, 3600, 2600],
    }),
    /must retain the source table total width/,
  );
  const nonCanonicalTableWidthSourcePath = path.join(outputDir, "source-bound-table-column-widths-noncanonical-source.docx");
  const nonCanonicalTableWidthZip = await JSZip.loadAsync(tableWidthSourceBytes);
  nonCanonicalTableWidthZip.file("word/document.xml", tableWidthSourceXml.replace(/<w:gridCol w:w="2100"/, '<w:gridCol w:w="02100"'));
  await fs.writeFile(nonCanonicalTableWidthSourcePath, await nonCanonicalTableWidthZip.generateAsync({ type: "nodebuffer" }), { flag: "wx" });
  await assert.rejects(
    () => editImportedTableColumnWidths({
      inputPath: nonCanonicalTableWidthSourcePath,
      outputPath: path.join(outputDir, "source-bound-table-column-widths-noncanonical.docx"),
      auditPath: path.join(outputDir, "source-bound-table-column-widths-noncanonical-audit.json"),
      tableBlockIndex: 1,
      expectedColumnWidthsDxa: sourceColumnWidths,
      replacementColumnWidthsDxa: replacementColumnWidths,
    }),
    /canonical unsigned integer/,
  );
  assert.deepEqual(await fs.readFile(tableWidthSourcePath), tableWidthSourceBytes);

  await assert.rejects(
    () => editImportedSectionMargins({
      inputPath: marginWorkflowSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-margins-mismatched.docx"),
      auditPath: path.join(outputDir, "source-bound-section-margins-mismatched-audit.json"),
      sectionBlockIndex: 1,
      expectedMargins: { ...sourceMargins, left: 999 },
      replacementMargins,
    }),
    /margins do not match the expected source value/,
  );
  await assert.rejects(
    () => editImportedSectionMargins({
      inputPath: marginWorkflowSourcePath,
      outputPath: path.join(outputDir, "source-bound-section-margins-invalid.docx"),
      auditPath: path.join(outputDir, "source-bound-section-margins-invalid-audit.json"),
      sectionBlockIndex: 1,
      expectedMargins: sourceMargins,
      replacementMargins: { top: 1440, right: 1440, bottom: 1440, left: 1728 },
    }),
    /replacementMargins\.gutter is required/,
  );
  await assert.rejects(
    () => editImportedSectionMargins({
      inputPath: marginWorkflowSourcePath,
      outputPath: marginWorkflowOutputPath,
      auditPath: path.join(outputDir, "source-bound-section-margins-overwrite-audit.json"),
      sectionBlockIndex: 1,
      expectedMargins: sourceMargins,
      replacementMargins,
    }),
    /outputPath already exists; refusing to overwrite it/,
  );
  assert.deepEqual(await fs.readFile(marginWorkflowSourcePath), marginWorkflowSourceBytes);

  const toc = await runFixture("office-kit-toc", {
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  const tocDocument = await DocumentFile.importDocx(await FileBlob.load(toc.docxPath));
  const tocField = tocDocument.blocks.find((block) => block.kind === "field");
  assert.equal(tocDocument.settings.updateFields, true);
  assert.equal(tocField?.complex, true);
  assert.equal(tocField?.instruction, 'TOC \\o "1-4" \\h \\z \\u');
  assert.equal(tocField?.display, "Update this table of contents in Word");
  assert.equal(toc.qa.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  const tocZip = await JSZip.loadAsync(await fs.readFile(toc.docxPath));
  const tocXml = await tocZip.file("word/document.xml").async("text");
  const tocSettings = await tocZip.file("word/settings.xml").async("text");
  assert.match(tocXml, /w:fldCharType="begin"/);
  assert.match(tocXml, /<w:instrText[^>]*> TOC \\o (?:"|&quot;)1-4(?:"|&quot;) \\h \\z \\u <\/w:instrText>/);
  assert.match(tocXml, /w:fldCharType="separate"/);
  assert.match(tocXml, /w:fldCharType="end"/);
  assert.match(tocSettings, /<w:updateFields\b[^>]*w:val="true"/);

  const inlineFields = await runFixture("office-kit-inline-fields", {
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  const inlineFieldDocument = await DocumentFile.importDocx(await FileBlob.load(inlineFields.docxPath));
  const inlineFieldParagraph = inlineFieldDocument.blocks.find((block) => block.name === "field-caption" || block.text.startsWith("Figure 1:"));
  assert.equal(inlineFieldParagraph?.text, "Figure 1: Updated revenue. See figure 1 on page 1.");
  assert.deepEqual(inlineFieldParagraph?.runs.filter((run) => run.inlineField).map((run) => run.inlineField.instruction), [
    "SEQ Figure \\* ARABIC",
    "REF fig1 \\h",
    "PAGEREF fig1 \\h",
  ]);
  assert.equal(inlineFieldParagraph?.runs.find((run) => run.inlineField?.instruction.startsWith("SEQ "))?.inlineField.bookmarkName, "fig1");
  assert.equal(inlineFieldParagraph?.runs.find((run) => run.inlineField?.instruction.startsWith("SEQ "))?.inlineField.bookmarkNativeId, 0);
  assert.equal(inlineFields.qa.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  const inlineFieldZip = await JSZip.loadAsync(await fs.readFile(inlineFields.docxPath));
  const inlineFieldXml = await inlineFieldZip.file("word/document.xml").async("text");
  assert.equal((inlineFieldXml.match(/w:fldCharType="begin"/g) || []).length, 3);
  assert.equal((inlineFieldXml.match(/w:fldCharType="separate"/g) || []).length, 3);
  assert.equal((inlineFieldXml.match(/w:fldCharType="end"/g) || []).length, 3);
  assert.match(inlineFieldXml, /SEQ Figure \\[*] ARABIC/);
  assert.match(inlineFieldXml, /REF fig1 \\h/);
  assert.match(inlineFieldXml, /PAGEREF fig1 \\h/);
  const inlineBookmarkStart = inlineFieldXml.indexOf('w:name="fig1"');
  const inlineBookmarkResult = inlineFieldXml.indexOf("<w:t>1</w:t>", inlineBookmarkStart);
  const inlineBookmarkEnd = inlineFieldXml.indexOf("<w:bookmarkEnd", inlineBookmarkStart);
  assert.ok(inlineBookmarkStart >= 0 && inlineBookmarkResult > inlineBookmarkStart && inlineBookmarkEnd > inlineBookmarkResult, "Caption-number bookmark must wrap only the SEQ cached-result run.");

  const classicFixture = await runFixture("package-comments");
  const classicDocument = await DocumentFile.importDocx(await FileBlob.load(classicFixture.docxPath));
  assert.equal(classicDocument.comments.length, 1);
  assert.equal(classicDocument.comments[0].author, "QA Lead");
  assert.equal(classicDocument.comments[0].text, "Decision paragraph confirmed.");
  const classicSourceBytes = await fs.readFile(classicFixture.docxPath);
  const classicTarget = classicDocument.blocks.find((block) => block.id === classicDocument.comments[0].targetId);
  assert.equal(classicTarget?.kind, "paragraph");
  const { editClassicComment } = await import(
    "../skills/documents/skills/documents/examples/officekit-classic-comment-edit-workflow.mjs"
  );
  const classicWorkflowOutput = path.join(outputDir, "classic-comment-updated.docx");
  const classicWorkflowAudit = path.join(outputDir, "classic-comment-audit.json");
  const classicWorkflow = await editClassicComment({
    inputPath: classicFixture.docxPath,
    outputPath: classicWorkflowOutput,
    auditPath: classicWorkflowAudit,
    anchorText: classicTarget.text,
    expectedCommentText: classicDocument.comments[0].text,
    replacementText: "Decision paragraph approved after QA.",
  });
  assert.equal(classicWorkflow.audit.provider.actual, "office-kit");
  assert.equal(classicWorkflow.audit.validation.reimport.ok, true);
  assert.equal(classicWorkflow.audit.validation.modelRender.renderer, "model-svg");
  assert.deepEqual(await fs.readFile(classicFixture.docxPath), classicSourceBytes);
  const classicWorkflowDocument = await DocumentFile.importDocx(await FileBlob.load(classicWorkflowOutput));
  assert.equal(classicWorkflowDocument.comments.length, 1);
  assert.equal(classicWorkflowDocument.comments[0].id, classicDocument.comments[0].id);
  assert.equal(classicWorkflowDocument.comments[0].targetId, classicDocument.comments[0].targetId);
  assert.equal(classicWorkflowDocument.comments[0].author, "QA Lead");
  assert.equal(classicWorkflowDocument.comments[0].text, "Decision paragraph approved after QA.");

  const modernSourceDocument = DocumentModel.create({ name: "Modern review thread", blocks: [] });
  const modernTarget = modernSourceDocument.addParagraph("Decision: ship the bounded modern review thread.");
  const modernRoot = modernSourceDocument.addComment(modernTarget, "Please confirm the release evidence.", {
    author: "Lead reviewer",
    initials: "LR",
    date: "2026-07-19T08:00:00Z",
    resolved: false,
    paraId: "11111111",
    durableId: "33333333",
    dateUtc: "2026-07-19T08:00:00Z",
    person: { providerId: "provider-a", userId: "lead@example.test" },
  });
  modernSourceDocument.replyToComment(modernRoot, "The evidence is attached.", {
    author: "Release reviewer",
    initials: "RR",
    date: "2026-07-19T08:05:00Z",
    paraId: "22222222",
    durableId: "44444444",
    dateUtc: "2026-07-19T08:05:00Z",
    person: { providerId: "provider-b", userId: "release@example.test" },
  });
  const modernSourcePath = path.join(outputDir, "modern-comment-source.docx");
  await (await DocumentFile.exportDocx(modernSourceDocument)).save(modernSourcePath);
  const modernSourceBytes = await fs.readFile(modernSourcePath);
  const { editModernCommentThread } = await import(
    "../skills/documents/skills/documents/examples/officekit-modern-comment-thread-workflow.mjs"
  );
  const modernWorkflowOutput = path.join(outputDir, "modern-comment-reviewed.docx");
  const modernWorkflowAudit = path.join(outputDir, "modern-comment-audit.json");
  const modernWorkflow = await editModernCommentThread({
    inputPath: modernSourcePath,
    outputPath: modernWorkflowOutput,
    auditPath: modernWorkflowAudit,
    anchorText: "bounded modern review thread",
    expectedRootText: "Please confirm the release evidence.",
    replacementRootText: "Release evidence approved.",
    expectedReplyText: "The evidence is attached.",
    replacementReplyText: "Evidence retained with the approval.",
    resolved: true,
  });
  assert.equal(modernWorkflow.audit.provider.actual, "office-kit");
  assert.equal(modernWorkflow.audit.operation.resolved, true);
  assert.equal(modernWorkflow.audit.validation.reimport.commentCount, 2);
  assert.deepEqual(await fs.readFile(modernSourcePath), modernSourceBytes);
  const modernWorkflowDocument = await DocumentFile.importDocx(await FileBlob.load(modernWorkflowOutput));
  assert.deepEqual(modernWorkflowDocument.comments.map((comment) => [comment.text, comment.resolved]), [
    ["Release evidence approved.", true],
    ["Evidence retained with the approval.", false],
  ]);
  assert.equal(modernWorkflowDocument.comments[1].parentId, modernWorkflowDocument.comments[0].id);
  assert.equal(modernWorkflowDocument.comments[0].durableId, "33333333");
  assert.deepEqual(modernWorkflowDocument.comments[1].person, {
    providerId: "provider-b",
    userId: "release@example.test",
  });
  const modernZip = await JSZip.loadAsync(await fs.readFile(modernWorkflowOutput));
  for (const part of [
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsIds.xml",
    "word/commentsExtensible.xml",
    "word/people.xml",
  ]) assert.ok(modernZip.file(part), `Expected ${part}`);
  const modernDocumentXml = await modernZip.file("word/document.xml").async("text");
  assert.equal((modernDocumentXml.match(/<w:commentRangeStart\b/g) || []).length, 1);
  assert.equal((modernDocumentXml.match(/<w:commentRangeEnd\b/g) || []).length, 1);
  assert.equal((modernDocumentXml.match(/<w:commentReference\b/g) || []).length, 1);
  const modernRender = await verifyDocumentFile(modernWorkflowOutput, {
    outputDir: path.join(outputDir, "modern-comment-render"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(modernRender.summary.verifyOk, true);
  assert.equal(modernRender.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const fragmentedPatchDocument = DocumentModel.create({ name: "Fragmented source text patch", blocks: [] });
  fragmentedPatchDocument.addParagraph("Quarterly plan");
  fragmentedPatchDocument.addParagraph("Unchanged review context.");
  fragmentedPatchDocument.addTable({ values: [["Revenue", "42"]] });
  const fragmentedPatchAuthored = await DocumentFile.exportDocx(fragmentedPatchDocument);
  const fragmentedPatchZip = await JSZip.loadAsync(await fragmentedPatchAuthored.arrayBuffer());
  const fragmentedPatchOriginalXml = await fragmentedPatchZip.file("word/document.xml").async("text");
  const fragmentedPatchSourceXml = fragmentedPatchOriginalXml.replace(
    '<w:pPr><w:pStyle w:val="Normal" /></w:pPr><w:r><w:t>Quarterly plan</w:t></w:r>',
    '<w:pPr><w:pStyle w:val="Normal" /><w:widowControl /><w:suppressAutoHyphens /></w:pPr><w:r><w:t>Quarter</w:t></w:r><w:r><w:t>ly plan</w:t></w:r>',
  ).replace(
    '<w:r><w:rPr><w:b /></w:rPr><w:t>Revenue</w:t></w:r>',
    '<w:r><w:rPr><w:b /></w:rPr><w:t>Rev</w:t></w:r><w:r><w:rPr><w:b /></w:rPr><w:t>enue</w:t></w:r>',
  );
  assert.notEqual(fragmentedPatchSourceXml, fragmentedPatchOriginalXml);
  const fragmentedPatchSourcePath = path.join(outputDir, "fragmented-patch-source.docx");
  await (await DocumentFile.patchDocx(fragmentedPatchAuthored, [
    { path: "word/document.xml", xml: fragmentedPatchSourceXml },
  ])).save(fragmentedPatchSourcePath);
  const fragmentedPatchSourceBytes = await fs.readFile(fragmentedPatchSourcePath);
  const fragmentedPatchImported = await DocumentFile.importDocx(await FileBlob.load(fragmentedPatchSourcePath));
  const fragmentedPatchTargetIndex = fragmentedPatchImported.blocks.findIndex((block) => block.text === "Quarterly plan");
  assert.ok(fragmentedPatchTargetIndex >= 0);
  assert.equal(fragmentedPatchImported.blocks[fragmentedPatchTargetIndex].textEditable, false);
  assert.equal(fragmentedPatchImported.blocks[fragmentedPatchTargetIndex].textPatchable, true);
  const { patchImportedText } = await import(
    "../skills/documents/skills/documents/examples/officekit-source-text-patch-workflow.mjs"
  );
  const fragmentedPatchOutputPath = path.join(outputDir, "fragmented-patch-output.docx");
  const fragmentedPatchAuditPath = path.join(outputDir, "fragmented-patch-audit.json");
  const fragmentedPatchWorkflow = await patchImportedText({
    inputPath: fragmentedPatchSourcePath,
    outputPath: fragmentedPatchOutputPath,
    auditPath: fragmentedPatchAuditPath,
    target: { kind: "paragraph", blockIndex: fragmentedPatchTargetIndex },
    search: "Quarterly",
    replacement: "Annual",
  });
  assert.equal(fragmentedPatchWorkflow.audit.provider.actual, "office-kit");
  assert.equal(fragmentedPatchWorkflow.audit.provider.silentFallback, false);
  assert.deepEqual(fragmentedPatchWorkflow.audit.validation.changedParts, ["word/document.xml"]);
  assert.equal(fragmentedPatchWorkflow.audit.validation.reimport.textPatchable, true);
  assert.deepEqual(await fs.readFile(fragmentedPatchSourcePath), fragmentedPatchSourceBytes);
  const fragmentedPatchRoundTrip = await DocumentFile.importDocx(await FileBlob.load(fragmentedPatchOutputPath));
  assert.equal(fragmentedPatchRoundTrip.blocks[fragmentedPatchTargetIndex].text, "Annual plan");
  assert.equal(fragmentedPatchRoundTrip.blocks[1].text, "Unchanged review context.");
  const fragmentedPatchTableIndex = fragmentedPatchRoundTrip.blocks.findIndex((block) => block.kind === "table");
  assert.ok(fragmentedPatchTableIndex >= 0);
  assert.equal(fragmentedPatchRoundTrip.blocks[fragmentedPatchTableIndex].getCell(0, 0).textPatchable, true);
  const fragmentedPatchTableSourceBytes = await fs.readFile(fragmentedPatchOutputPath);
  const fragmentedPatchFinalPath = path.join(outputDir, "fragmented-patch-table-output.docx");
  const fragmentedPatchTableAuditPath = path.join(outputDir, "fragmented-patch-table-audit.json");
  const fragmentedPatchTableWorkflow = await patchImportedText({
    inputPath: fragmentedPatchOutputPath,
    outputPath: fragmentedPatchFinalPath,
    auditPath: fragmentedPatchTableAuditPath,
    target: { kind: "tableCell", blockIndex: fragmentedPatchTableIndex, row: 0, column: 0 },
    search: "Revenue",
    replacement: "Net revenue",
  });
  assert.equal(fragmentedPatchTableWorkflow.audit.operation.target.kind, "tableCell");
  assert.deepEqual(fragmentedPatchTableWorkflow.audit.validation.changedParts, ["word/document.xml"]);
  assert.deepEqual(await fs.readFile(fragmentedPatchOutputPath), fragmentedPatchTableSourceBytes);
  const fragmentedPatchFinal = await DocumentFile.importDocx(await FileBlob.load(fragmentedPatchFinalPath));
  assert.equal(fragmentedPatchFinal.blocks[fragmentedPatchTargetIndex].text, "Annual plan");
  assert.equal(fragmentedPatchFinal.blocks[fragmentedPatchTableIndex].getCell(0, 0).value, "Net revenue");
  const fragmentedPatchRender = await verifyDocumentFile(fragmentedPatchFinalPath, {
    outputDir: path.join(outputDir, "fragmented-patch-render"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(fragmentedPatchRender.summary.verifyOk, true);
  assert.equal(fragmentedPatchRender.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const trackedReplacementSourceDocument = DocumentModel.create({
    name: "Tracked replacement source",
    blocks: [],
  });
  trackedReplacementSourceDocument.addParagraph("The draft budget assumes 30 days of cash buffer.", {
    runs: [
      { text: "The draft budget assumes 3", style: { bold: true, color: "#315A83" } },
      { text: "0 da", style: { bold: true, color: "#315A83" } },
      { text: "ys of cash buffer.", style: { bold: true, color: "#315A83" } },
    ],
  });
  trackedReplacementSourceDocument.addParagraph("Unchanged review context.");
  const trackedReplacementSourcePath = path.join(outputDir, "tracked-replacement-source.docx");
  await (await DocumentFile.exportDocx(trackedReplacementSourceDocument)).save(trackedReplacementSourcePath);
  const trackedReplacementSourceBytes = await fs.readFile(trackedReplacementSourcePath);
  const { addDocumentTrackedReplacement } = await import(
    "../skills/documents/skills/documents/examples/officekit-tracked-replacement-workflow.mjs"
  );
  const trackedReplacementPath = path.join(outputDir, "tracked-replacement.docx");
  const trackedReplacementAuditPath = path.join(outputDir, "tracked-replacement-audit.json");
  const trackedReplacementWorkflow = await addDocumentTrackedReplacement({
    inputPath: trackedReplacementSourcePath,
    outputPath: trackedReplacementPath,
    auditPath: trackedReplacementAuditPath,
    expectedText: "The draft budget assumes 30 days of cash buffer.",
    search: "30 days",
    replacement: "45 days",
    author: "Budget reviewer",
    date: "2026-07-21T09:30:00Z",
  });
  assert.equal(trackedReplacementWorkflow.audit.provider.actual, "office-kit");
  assert.equal(trackedReplacementWorkflow.audit.provider.silentFallback, false);
  assert.equal(trackedReplacementWorkflow.audit.savePolicy.overwrite, false);
  assert.deepEqual(trackedReplacementWorkflow.audit.operation.changedParts, ["word/document.xml"]);
  assert.equal(trackedReplacementWorkflow.audit.operation.targetBlockIndex, 0);
  assert.equal(trackedReplacementWorkflow.audit.operation.matchedSourceRunCount, 3);
  assert.deepEqual(await fs.readFile(trackedReplacementSourcePath), trackedReplacementSourceBytes);
  const trackedReplacementZip = await JSZip.loadAsync(await fs.readFile(trackedReplacementPath));
  const trackedReplacementXml = await trackedReplacementZip.file("word/document.xml").async("text");
  assert.equal((trackedReplacementXml.match(/<w:del\b/g) || []).length, 1);
  assert.equal((trackedReplacementXml.match(/<w:ins\b/g) || []).length, 1);
  assert.deepEqual(
    [...trackedReplacementXml.matchAll(/<w:delText(?:\s[^>]*)?>([\s\S]*?)<\/w:delText>/g)].map((match) => match[1]),
    ["3", "0 da", "ys"],
  );
  assert.match(trackedReplacementXml, /<w:t>45 days<\/w:t>/);
  const trackedReplacementDocument = await DocumentFile.importDocx(await FileBlob.load(trackedReplacementPath));
  assert.equal(trackedReplacementDocument.blocks[0].text, "The draft budget assumes 45 days of cash buffer.");
  assert.equal(trackedReplacementDocument.blocks[0].textEditable, false);
  const trackedReplacementRender = await verifyDocumentFile(trackedReplacementPath, {
    outputDir: path.join(outputDir, "tracked-replacement-render"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(trackedReplacementRender.summary.verifyOk, true);
  assert.equal(trackedReplacementRender.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const trackedReplacementBytes = await fs.readFile(trackedReplacementPath);
  const trackedReplacementSha256 = createHash("sha256").update(trackedReplacementBytes).digest("hex");
  const acceptedTrackedReplacement = await DocumentFile.finalizeRevisions(new FileBlob(trackedReplacementBytes), {
    mode: "accept",
    expectedSourceSha256: trackedReplacementSha256,
  });
  assert.equal(acceptedTrackedReplacement.metadata.revisionFinalization.insertionCount, 1);
  assert.equal(acceptedTrackedReplacement.metadata.revisionFinalization.deletionCount, 1);
  const acceptedTrackedReplacementPath = path.join(outputDir, "tracked-replacement-accepted.docx");
  await acceptedTrackedReplacement.save(acceptedTrackedReplacementPath);
  const acceptedTrackedReplacementDocument = await DocumentFile.importDocx(acceptedTrackedReplacement);
  assert.equal(acceptedTrackedReplacementDocument.blocks[0].text, "The draft budget assumes 45 days of cash buffer.");
  const acceptedTrackedReplacementRender = await verifyDocumentFile(acceptedTrackedReplacementPath, {
    outputDir: path.join(outputDir, "tracked-replacement-accepted-render"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(acceptedTrackedReplacementRender.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const rejectedTrackedReplacement = await DocumentFile.finalizeRevisions(new FileBlob(trackedReplacementBytes), {
    mode: "reject",
    expectedSourceSha256: trackedReplacementSha256,
  });
  const rejectedTrackedReplacementPath = path.join(outputDir, "tracked-replacement-rejected.docx");
  await rejectedTrackedReplacement.save(rejectedTrackedReplacementPath);
  const rejectedTrackedReplacementDocument = await DocumentFile.importDocx(rejectedTrackedReplacement);
  assert.equal(rejectedTrackedReplacementDocument.blocks[0].text, "The draft budget assumes 30 days of cash buffer.");
  const rejectedTrackedReplacementRender = await verifyDocumentFile(rejectedTrackedReplacementPath, {
    outputDir: path.join(outputDir, "tracked-replacement-rejected-render"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(rejectedTrackedReplacementRender.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  await assert.rejects(
    () => addDocumentTrackedReplacement({
      inputPath: trackedReplacementSourcePath,
      outputPath: trackedReplacementPath,
      auditPath: path.join(outputDir, "must-not-publish-tracked-replacement-audit.json"),
      expectedText: "The draft budget assumes 30 days of cash buffer.",
      search: "30 days",
      replacement: "45 days",
      author: "Budget reviewer",
    }),
    (error) => error?.code === "EEXIST",
  );

  const tableTrackedSource = DocumentModel.create({ name: "Table tracked replacement source", blocks: [] });
  tableTrackedSource.addParagraph("Contract review matrix");
  tableTrackedSource.addTable({
    name: "contract-terms",
    styleId: "TableGrid",
    widthDxa: 9000,
    indentDxa: 120,
    columnWidthsDxa: [2800, 6200],
    cellMarginsDxa: { top: 80, right: 120, bottom: 80, left: 120 },
    borderColor: "445566",
    borderSize: 8,
    headerFill: "E2E8F0",
    values: [["Term", "Current wording"], ["Payment", "Payment is due in 30 days."]],
  });
  tableTrackedSource.addParagraph("Unchanged approval context.");
  const tableTrackedSourcePath = path.join(outputDir, "table-tracked-source.docx");
  await (await DocumentFile.exportDocx(tableTrackedSource)).save(tableTrackedSourcePath);
  const tableTrackedSourceBytes = await fs.readFile(tableTrackedSourcePath);
  const tableTrackedPath = path.join(outputDir, "table-tracked.docx");
  const tableTrackedAuditPath = path.join(outputDir, "table-tracked-audit.json");
  const tableTrackedWorkflow = await addDocumentTrackedReplacement({
    inputPath: tableTrackedSourcePath,
    outputPath: tableTrackedPath,
    auditPath: tableTrackedAuditPath,
    expectedText: "Payment is due in 30 days.",
    search: "30 days",
    replacement: "45 days",
    author: "Contract reviewer",
    date: "2026-07-21T11:00:00Z",
  });
  assert.deepEqual(tableTrackedWorkflow.audit.operation.target, {
    kind: "tableCell",
    blockIndex: 1,
    row: 1,
    column: 1,
  });
  assert.deepEqual(tableTrackedWorkflow.audit.operation.changedParts, ["word/document.xml"]);
  assert.deepEqual(await fs.readFile(tableTrackedSourcePath), tableTrackedSourceBytes);
  const tableTrackedDocument = await DocumentFile.importDocx(await FileBlob.load(tableTrackedPath));
  assert.equal(tableTrackedDocument.blocks[1].getCell(1, 1).value, "Payment is due in 45 days.");
  assert.equal(tableTrackedDocument.blocks[1].getCell(1, 1).editable, false);
  const tableTrackedRender = await verifyDocumentFile(tableTrackedPath, {
    outputDir: path.join(outputDir, "table-tracked-render"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(tableTrackedRender.summary.verifyOk, true);
  assert.equal(tableTrackedRender.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const tableTrackedBytes = await fs.readFile(tableTrackedPath);
  const tableTrackedSha256 = createHash("sha256").update(tableTrackedBytes).digest("hex");
  const acceptedTableTracked = await DocumentFile.finalizeRevisions(new FileBlob(tableTrackedBytes), {
    mode: "accept",
    expectedSourceSha256: tableTrackedSha256,
  });
  const rejectedTableTracked = await DocumentFile.finalizeRevisions(new FileBlob(tableTrackedBytes), {
    mode: "reject",
    expectedSourceSha256: tableTrackedSha256,
  });
  assert.equal((await DocumentFile.importDocx(acceptedTableTracked)).blocks[1].getCell(1, 1).value, "Payment is due in 45 days.");
  assert.equal((await DocumentFile.importDocx(rejectedTableTracked)).blocks[1].getCell(1, 1).value, "Payment is due in 30 days.");

  const revisionSourceDocument = DocumentModel.create({
    name: "Bounded revision finalization",
    settings: { trackRevisions: true },
    blocks: [],
  });
  revisionSourceDocument.addParagraph("Revision review baseline.");
  revisionSourceDocument.addInsertion("Accepted insertion.", {
    author: "Release reviewer",
    date: "2026-07-21T08:00:00Z",
  });
  revisionSourceDocument.addDeletion("Rejected legacy wording.", {
    author: "Release reviewer",
    date: "2026-07-21T08:05:00Z",
  });
  revisionSourceDocument.addParagraph("Revision review complete.");
  const revisionSourcePath = path.join(outputDir, "revision-source.docx");
  await (await DocumentFile.exportDocx(revisionSourceDocument)).save(revisionSourcePath);
  const revisionSourceBytes = await fs.readFile(revisionSourcePath);
  const { finalizeDocumentRevisions } = await import(
    "../skills/documents/skills/documents/examples/officekit-revision-finalization-workflow.mjs"
  );
  const acceptedRevisionPath = path.join(outputDir, "revision-accepted.docx");
  const acceptedRevisionAuditPath = path.join(outputDir, "revision-accepted-audit.json");
  const acceptedRevisionWorkflow = await finalizeDocumentRevisions({
    inputPath: revisionSourcePath,
    outputPath: acceptedRevisionPath,
    auditPath: acceptedRevisionAuditPath,
    mode: "accept",
  });
  assert.equal(acceptedRevisionWorkflow.audit.provider.actual, "office-kit");
  assert.equal(acceptedRevisionWorkflow.audit.provider.silentFallback, false);
  assert.equal(acceptedRevisionWorkflow.audit.savePolicy.overwrite, false);
  assert.deepEqual(acceptedRevisionWorkflow.audit.operation.changedParts, ["word/document.xml", "word/settings.xml"]);
  assert.equal(acceptedRevisionWorkflow.audit.validation.reimport.remainingRevisions, 0);
  assert.deepEqual(await fs.readFile(revisionSourcePath), revisionSourceBytes);
  const acceptedRevisionDocument = await DocumentFile.importDocx(await FileBlob.load(acceptedRevisionPath));
  assert.equal(acceptedRevisionDocument.settings.trackRevisions, false);
  assert.equal(acceptedRevisionDocument.blocks.some((block) => block.kind === "change"), false);
  assert.equal(acceptedRevisionDocument.blocks.some((block) => block.text === "Accepted insertion."), true);
  assert.equal(acceptedRevisionDocument.blocks.some((block) => block.text === "Rejected legacy wording."), false);
  const acceptedRevisionRender = await verifyDocumentFile(acceptedRevisionPath, {
    outputDir: path.join(outputDir, "revision-accepted-render"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(acceptedRevisionRender.summary.verifyOk, true);
  assert.equal(acceptedRevisionRender.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const rejectedRevisionPath = path.join(outputDir, "revision-rejected.docx");
  const rejectedRevisionAuditPath = path.join(outputDir, "revision-rejected-audit.json");
  const rejectedRevisionWorkflow = await finalizeDocumentRevisions({
    inputPath: revisionSourcePath,
    outputPath: rejectedRevisionPath,
    auditPath: rejectedRevisionAuditPath,
    mode: "reject",
    keepTracking: true,
  });
  assert.deepEqual(rejectedRevisionWorkflow.audit.operation.changedParts, ["word/document.xml"]);
  assert.equal(rejectedRevisionWorkflow.audit.operation.trackingAfter, true);
  const rejectedRevisionDocument = await DocumentFile.importDocx(await FileBlob.load(rejectedRevisionPath));
  assert.equal(rejectedRevisionDocument.settings.trackRevisions, true);
  assert.equal(rejectedRevisionDocument.blocks.some((block) => block.text === "Accepted insertion."), false);
  assert.equal(rejectedRevisionDocument.blocks.some((block) => block.text === "Rejected legacy wording."), true);
  const acceptedRevisionBytes = await fs.readFile(acceptedRevisionPath);
  await assert.rejects(
    () => finalizeDocumentRevisions({
      inputPath: revisionSourcePath,
      outputPath: acceptedRevisionPath,
      auditPath: path.join(outputDir, "must-not-publish-audit.json"),
      mode: "accept",
    }),
    (error) => error?.code === "EEXIST",
  );
  assert.deepEqual(await fs.readFile(acceptedRevisionPath), acceptedRevisionBytes);

  const directNumbering = await runFixture("package-numbering");
  const directNumberingDocument = await DocumentFile.importDocx(await FileBlob.load(directNumbering.docxPath));
  assert.equal(directNumberingDocument.blocks.filter((block) => block.kind === "listItem").length, 2);
  assert.equal(directNumberingDocument.blocks.some((block) => block.text === "Confirm the edited second item."), true);

  const sectionSettings = await runFixture("package-settings", {
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  const settingsDocument = await DocumentFile.importDocx(await FileBlob.load(sectionSettings.docxPath));
  assert.equal(settingsDocument.settings.evenAndOddHeaders, true);
  assert.equal(settingsDocument.settings.mirrorMargins, true);
  assert.equal(settingsDocument.settings.gutterAtTop, true);
  assert.equal(settingsDocument.blocks.find((block) => block.kind === "section")?.margins.gutter, 720);
  assert.deepEqual(settingsDocument.blocks.find((block) => block.kind === "section")?.lineNumbering, { countBy: 5, start: 0, distance: 360, restart: "newPage" });
  assert.deepEqual(settingsDocument.blocks.find((block) => block.kind === "section")?.pageNumbering, { start: 1, format: "lowerRoman" });
  assert.deepEqual(settingsDocument.blocks.find((block) => block.kind === "section")?.columns, { definitions: [{ width: 3000, spacing: 720 }, { width: 5640, spacing: 0 }], separator: true });
  assert.equal(settingsDocument.blocks.find((block) => block.text === "Review first-page and even-page variants.")?.paragraphFormat.suppressLineNumbers, true);
  assert.equal(settingsDocument.blocks.find((block) => block.text === "This paragraph remains in the section line-number sequence.")?.paragraphFormat.suppressLineNumbers, undefined);
  assert.equal(settingsDocument.sectionSettings[0]?.differentFirstPage, true);
  assert.equal(settingsDocument.headers.some((item) => item.referenceType === "first"), true);
  assert.equal(settingsDocument.headers.some((item) => item.referenceType === "even"), true);
  assert.equal(settingsDocument.footers[0]?.fieldInstruction, "PAGE");
  const packageSettingsZip = await JSZip.loadAsync(await fs.readFile(sectionSettings.docxPath));
  assert.match(await packageSettingsZip.file("word/settings.xml").async("text"), /<w:mirrorMargins\s*\/>/);
  assert.match(await packageSettingsZip.file("word/settings.xml").async("text"), /<w:gutterAtTop\s*\/>/);
  assert.match(await packageSettingsZip.file("word/document.xml").async("text"), /<w:pgMar\b(?=[^>]*w:gutter="720")[^>]*\/>/);
  assert.match(await packageSettingsZip.file("word/document.xml").async("text"), /<w:lnNumType\b(?=[^>]*w:countBy="5")(?=[^>]*w:start="0")(?=[^>]*w:distance="360")(?=[^>]*w:restart="newPage")[^>]*\/>/);
  assert.match(await packageSettingsZip.file("word/document.xml").async("text"), /<w:p>[\s\S]*?<w:suppressLineNumbers\b[^>]*w:val="true"[^>]*\/>[\s\S]*?Review first-page and even-page variants\.[\s\S]*?<\/w:p>/);
  assert.match(await packageSettingsZip.file("word/document.xml").async("text"), /<w:pgNumType\b(?=[^>]*w:start="1")(?=[^>]*w:fmt="lowerRoman")[^>]*\/>/);
  assert.match(await packageSettingsZip.file("word/document.xml").async("text"), /<w:cols\b(?=[^>]*w:equalWidth="(?:false|0)")(?=[^>]*w:sep="(?:true|1)")[^>]*>[\s\S]*?<w:col\b(?=[^>]*w:w="3000")(?=[^>]*w:space="720")[^>]*\/>[\s\S]*?<w:col\b(?=[^>]*w:w="5640")[^>]*\/>[\s\S]*?<\/w:cols>/);
  assert.equal(sectionSettings.qa.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  if (nativeStatus.available) assert.equal(sectionSettings.qa.summary.nativeRender.ok, true);

  const protection = await runFixture("office-kit-protection");
  const protectedDocument = await DocumentFile.importDocx(await FileBlob.load(protection.docxPath));
  assert.deepEqual(protectedDocument.settings.documentProtection, {
    edit: "comments",
    enforcement: true,
    formatting: false,
  });
  const protectionZip = await JSZip.loadAsync(await fs.readFile(protection.docxPath));
  assert.match(
    await protectionZip.file("word/settings.xml").async("text"),
    /<w:documentProtection(?=[^>]*w:edit="comments")(?=[^>]*w:enforcement="true")(?=[^>]*w:formatting="false")[^>]*\/>/,
  );

  const unprotectedFixture = structuredClone(protection.fixture);
  unprotectedFixture.settings.documentProtection = false;
  unprotectedFixture.edits = [];
  const unprotectedPath = path.join(outputDir, "office-kit-protection", "unprotected-layout-control.docx");
  await (await DocumentFile.exportDocx(createDocumentFromFixture(unprotectedFixture))).save(unprotectedPath);
  const protectionBaselineDir = path.join(outputDir, "protection-layout-baseline");
  await verifyDocumentFile(unprotectedPath, {
    outputDir: path.join(outputDir, "protection-unprotected-qa"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
    baselineDir: protectionBaselineDir,
    writeBaseline: true,
  });
  const protectedLayoutQa = await verifyDocumentFile(protection.docxPath, {
    outputDir: path.join(outputDir, "protection-protected-qa"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "auto",
    baselineDir: protectionBaselineDir,
  });
  assert.equal(protectedLayoutQa.summary.modelPixelDiff.changed, false);
  if (nativeStatus.available) {
    assert.equal(protectedLayoutQa.summary.nativeRender.pageCountMatches, true);
    assert.equal(protectedLayoutQa.summary.nativeRender.pages.every((page) => page.pixelDiff.changed === false), true);
  }

  const baselineWrite = await verifyDocumentFile(business.docxPath, {
    outputDir: path.join(outputDir, "baseline-write"),
    previewFormat: "png",
    nativeRender: "off",
    baselineDir,
    writeBaseline: true,
  });
  assert.equal(baselineWrite.summary.writeBaseline, true);
  assert.ok((await fs.stat(baselineWrite.summary.modelBaselinePath)).size > 100);
  const baselineCompare = await verifyDocumentFile(business.docxPath, {
    outputDir: path.join(outputDir, "baseline-compare"),
    previewFormat: "png",
    nativeRender: "off",
    baselineDir,
  });
  assert.equal(baselineCompare.summary.modelBaselineCompared, true);
  assert.equal(baselineCompare.summary.modelPixelDiff.changed, false);
  assert.equal(baselineCompare.summary.visualQaOk, true);

  const packageJson = JSON.parse(await fs.readFile(path.join(repoRoot, "package.json"), "utf8"));
  for (const shippedSkillPath of [
    "skills/documents/**",
    "skills/spreadsheets/**",
    "skills/presentations/**",
    "skills/pdf/**",
    "skills/template-creator/**",
    "skills/default-template-library/**",
  ]) {
    assert.ok(packageJson.files.includes(shippedSkillPath));
  }
  assert.ok(!packageJson.files.includes("skills/**"));
  assert.ok(packageJson.files.includes("skills/default-template-library/**"));
  const skillText = await fs.readFile(path.join(repoRoot, "skills", "documents", "skills", "documents", "SKILL.md"), "utf8");
  const createEditGuide = await fs.readFile(path.join(repoRoot, "skills", "documents", "skills", "documents", "tasks", "create_edit.md"), "utf8");
  const templateCreateGuide = await fs.readFile(path.join(repoRoot, "skills", "documents", "skills", "documents", "template-create.md"), "utf8");
  const pluginReadme = await fs.readFile(path.join(repoRoot, "skills", "documents", "README.md"), "utf8");
  assert.match(pluginReadme, /office-kit/);
  assert.match(pluginReadme, /OfficeKit/);
  assert.match(skillText, /render_docx\.py/);
  assert.match(skillText, /DocumentModel/);
  assert.match(skillText, /DocumentFile/);
  assert.match(skillText, /OfficeKit/);
  assert.match(skillText, /artifact_tool\/API_QUICK_START\.md/);
  assert.match(skillText, /Classify multiple references by deliverable/);
  assert.match(skillText, /font\s+family, font size, color, paragraph spacing, line spacing, and margins/);
  assert.match(skillText, /stop while a controlling value remains[\s\S]*unresolved/);
  assert.match(skillText, /Preserve punctuation and whitespace in user text, quotations, templates, and[\s\S]*source listings/);
  assert.doesNotMatch(skillText, /Prefer ASCII punctuation/);
  assert.match(createEditGuide, /If the user supplied a DOCX, format sample, or explicit formatting rules[\s\S]*stop this ordinary create flow/);
  assert.ok(createEditGuide.toLowerCase().indexOf("template-distill.md") < createEditGuide.toLowerCase().indexOf("choose a design preset"));
  assert.match(createEditGuide, /no controlling value remains unresolved/);
  assert.match(createEditGuide, /With no controlling reference or explicit format specification/);
  assert.match(templateCreateGuide, /stop if its page geometry, font family,[\s\S]*margins remain unresolved/);
  assert.match(templateCreateGuide, /Preserve exact characters and whitespace inside user text,[\s\S]*source listings/);
  assert.match(skillText, /document\.addInsertion/);
  assert.match(skillText, /document\.addDeletion/);
  assert.match(skillText, /paragraph\.addTextContentControl/);
  assert.match(skillText, /document\.addBlockTextContentControl/);
  assert.match(skillText, /table\.getCell\(row, column\)\.addTextContentControl/);
  assert.match(skillText, /table\.getCell\(row, column\)[\s\S]*addCheckboxContentControl/);
  assert.match(skillText, /table\.getCell\(row, column\)[\s\S]*addDropdownContentControl/);
  assert.match(skillText, /table\.getCell\(row, column\)[\s\S]*addComboBoxContentControl/);
  assert.match(skillText, /table\.getCell\(row, column\)[\s\S]*addDateContentControl/);
  assert.match(skillText, /paragraph\.addCheckboxContentControl/);
  assert.match(skillText, /paragraph\.addDropdownContentControl/);
  assert.match(skillText, /paragraph\.addComboBoxContentControl/);
  assert.match(skillText, /paragraph\.addDateContentControl/);
  assert.match(skillText, /document\.fillContentControls/);
  assert.match(skillText, /document\.setCheckboxContentControls/);
  assert.match(skillText, /document\.setDropdownContentControls/);
  assert.match(skillText, /document\.setComboBoxContentControls/);
  assert.match(skillText, /document\.setDateContentControls/);
  assert.match(skillText, /document\.setSettings\(\{ documentProtection/);
  assert.match(skillText, /document\.setSettings\(\{ mirrorMargins: true \}\)/);
  assert.match(skillText, /document\.setSettings\(\{ gutterAtTop: true \}\)/);
  assert.match(skillText, /columns: \{ count: 2, spacing: 720, separator: true \}/);
  assert.match(skillText, /definitions: \[\s*\{ width: 3000, spacing: 720 \},\s*\{ width: 5640, spacing: 0 \}/);
  assert.match(skillText, /pageNumbering: \{ start: 1, format: "lowerRoman" \}/);
  assert.match(skillText, /lineNumbering: \{ countBy: 5, start: 0, distance: 360, restart: "newPage" \}/);
  assert.match(skillText, /paragraphFormat: \{ suppressLineNumbers: true \}/);
  assert.match(skillText, /paragraphFormat: \{ shadingFill: "#FEF3C7" \}/);
  assert.match(skillText, /paragraphFormat: \{ borders: \{ bottom: \{ color: "#315A83", size: 8, space: 2 \} \} \}/);
  assert.match(skillText, /verticalAlignment: "center"/);
  assert.match(skillText, /document\.addBibliographySource/);
  assert.match(skillText, /document\.addBibliography/);
  assert.match(skillText, /document\.addCitation/);
  assert.match(skillText, /document\.addTableOfContents/);
  assert.match(skillText, /paragraph\.addField/);
  assert.match(skillText, /officekit-source-text-patch-workflow\.mjs/);
  assert.match(skillText, /officekit-classic-comment-edit-workflow\.mjs/);
  assert.match(skillText, /officekit-header-text-edit-workflow\.mjs/);
  assert.match(skillText, /officekit-footer-text-edit-workflow\.mjs/);
  assert.match(skillText, /officekit-section-margin-edit-workflow\.mjs/);
  assert.match(skillText, /officekit-note-text-edit-workflow\.mjs/);
  assert.match(skillText, /officekit-modern-comment-thread-workflow\.mjs/);
  assert.match(skillText, /officekit-watermark-workflow\.mjs/);
  assert.match(skillText, /tasks\/headers_footers\.md/);
  assert.match(skillText, /officekit-tracked-replacement-workflow\.mjs/);
  assert.match(skillText, /officekit-revision-finalization-workflow\.mjs/);
  assert.doesNotMatch(skillText, /Author\/edit with `python-docx`|Default tool: python-docx/);
  const commentsGuide = await fs.readFile(path.join(repoRoot, "skills", "documents", "skills", "documents", "tasks", "comments_manage.md"), "utf8");
  assert.match(commentsGuide, /document\.addComment/);
  assert.match(commentsGuide, /document\.replyToComment/);
  assert.match(commentsGuide, /\.resolve\(\)/);
  assert.doesNotMatch(commentsGuide, /If the task is to \*insert\* new comments.+use the OOXML-level guide/);
  const manifestText = await fs.readFile(path.join(repoRoot, "skills", "documents", "skills", "documents", "manifest.txt"), "utf8");
  assert.match(manifestText, /^examples\/officekit-source-text-patch-workflow\.mjs$/m);
  assert.match(manifestText, /^examples\/officekit-page-furniture-text-edit\.mjs$/m);
  assert.match(manifestText, /^examples\/officekit-header-text-edit-workflow\.mjs$/m);
  assert.match(manifestText, /^examples\/officekit-footer-text-edit-workflow\.mjs$/m);
  assert.match(manifestText, /^examples\/officekit-section-margin-edit-workflow\.mjs$/m);
  assert.match(manifestText, /^examples\/officekit-note-text-edit-workflow\.mjs$/m);
  assert.match(manifestText, /^artifact_tool\/_source_bound_docx\.mjs$/m);
  assert.match(manifestText, /^artifact_tool\/_source_bound_sections\.mjs$/m);
  assert.match(manifestText, /^examples\/officekit-modern-comment-thread-workflow\.mjs$/m);
  assert.match(manifestText, /^examples\/officekit-watermark-workflow\.mjs$/m);
  assert.match(manifestText, /^tasks\/headers_footers\.md$/m);
  assert.match(manifestText, /^examples\/officekit-tracked-replacement-workflow\.mjs$/m);
  assert.match(manifestText, /^examples\/officekit-revision-finalization-workflow\.mjs$/m);
  assert.match(manifestText, /^examples\/end_to_end_smoke_test\.md$/m);
  const watermarkGuide = await fs.readFile(path.join(repoRoot, "skills", "documents", "skills", "documents", "tasks", "watermarks_background.md"), "utf8");
  assert.match(watermarkGuide, /document\.addWatermark/);
  assert.match(watermarkGuide, /watermark\.remove\(\)/);
  assert.match(watermarkGuide, /officekit-watermark-workflow\.mjs/);
  assert.match(watermarkGuide, /exactly one `word\/headerN\.xml`/);
  assert.match(watermarkGuide, /shared header.*multiple/is);
  assert.match(watermarkGuide, /image watermarks.*DrawingML watermarks.*irregular VML/is);
  const headersFootersGuide = await fs.readFile(path.join(repoRoot, "skills", "documents", "skills", "documents", "tasks", "headers_footers.md"), "utf8");
  assert.match(headersFootersGuide, /sourceBound.*editable/is);
  assert.match(headersFootersGuide, /at most one text edit.*part/is);
  assert.match(headersFootersGuide, /PAGE or other simple fields/);
  assert.match(headersFootersGuide, /officekit-header-text-edit-workflow\.mjs/);
  assert.match(headersFootersGuide, /officekit-footer-text-edit-workflow\.mjs/);
  assert.match(headersFootersGuide, /two entry points are intentionally separate/i);
  assert.match(headersFootersGuide, /fail closed/);
  const notesGuide = await fs.readFile(path.join(repoRoot, "skills", "documents", "skills", "documents", "tasks", "footnotes_endnotes.md"), "utf8");
  assert.match(notesGuide, /officekit-note-text-edit-workflow\.mjs/);
  assert.match(notesGuide, /word\/footnotes\.xml.*word\/endnotes\.xml/is);
  assert.match(notesGuide, /fail(?:s)? closed/i);
  const controlsGuide = await fs.readFile(path.join(repoRoot, "skills", "documents", "skills", "documents", "tasks", "forms_content_controls.md"), "utf8");
  assert.match(controlsGuide, /paragraph\.addTextContentControl/);
  assert.match(controlsGuide, /document\.addBlockTextContentControl/);
  assert.match(controlsGuide, /table\.getCell\(row, column\)\.addTextContentControl/);
  assert.match(controlsGuide, /paragraph\.addCheckboxContentControl/);
  assert.match(controlsGuide, /paragraph\.addDropdownContentControl/);
  assert.match(controlsGuide, /paragraph\.addComboBoxContentControl/);
  assert.match(controlsGuide, /paragraph\.addDateContentControl/);
  assert.match(controlsGuide, /document\.fillContentControls/);
  assert.match(controlsGuide, /document\.setCheckboxContentControls/);
  assert.match(controlsGuide, /document\.setDropdownContentControls/);
  assert.match(controlsGuide, /document\.setComboBoxContentControls/);
  assert.match(controlsGuide, /document\.setDateContentControls/);
  assert.match(controlsGuide, /rich.*multi-paragraph.*inline-within-cell.*nested.*data-bound.*irregular.*localized.*custom-symbol checkbox/is);
  const protectionGuide = await fs.readFile(path.join(repoRoot, "skills", "documents", "skills", "documents", "tasks", "protection_restrict_editing.md"), "utf8");
  assert.match(protectionGuide, /document\.setSettings\(\{ documentProtection/);
  assert.match(protectionGuide, /not encryption/i);
  assert.match(protectionGuide, /fail closed/i);
} finally {
  await fs.rm(outputDir, { recursive: true, force: true });
}

console.log("document skill smoke ok");
