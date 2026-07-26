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
const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-image-alt-text-workflow-"));
const nativeStatus = nativeDocumentRenderStatus();
const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2nGQAAAAASUVORK5CYII=";

function decodeXml(value) {
  return String(value).replace(/&(amp|lt|gt|quot|apos);/g, (_match, entity) => ({
    amp: "&",
    lt: "<",
    gt: ">",
    quot: "\"",
    apos: "'",
  })[entity]);
}

function rawImageDescriptions(xml) {
  return [...String(xml).matchAll(/<wp:docPr\b[^>]*\bdescr="([^"]*)"[^>]*\/>[\s\S]*?<pic:cNvPr\b[^>]*\bdescr="([^"]*)"[^>]*\/>/g)]
    .map((match) => ({ docPr: decodeXml(match[1]), nonVisual: decodeXml(match[2]) }));
}

function imageProjection(document) {
  return document.blocks.flatMap((block, blockIndex) => block.kind === "image" ? [{
    id: block.id,
    blockIndex,
    alt: block.alt,
    widthPx: block.widthPx,
    heightPx: block.heightPx,
    placement: block.placement ? structuredClone(block.placement) : undefined,
    dataUrlHash: createHash("sha256").update(block.dataUrl).digest("hex"),
  }] : []);
}

try {
  const sourceDocument = DocumentModel.create({ name: "Source-bound image alternative text", blocks: [] });
  sourceDocument.addParagraph("This report contains two floating visual markers.");
  sourceDocument.addImage({
    name: "target-image",
    dataUrl: PNG,
    alt: "Existing & approved quarterly chart",
    widthPx: 48,
    heightPx: 36,
    placement: {
      type: "floating",
      horizontal: { relativeTo: "margin", offsetPx: 36 },
      vertical: { relativeTo: "paragraph", offsetPx: 6 },
      wrap: "square",
      wrapSide: "bothSides",
      distanceFromTextPx: { top: 0, right: 4, bottom: 0, left: 4 },
    },
  });
  sourceDocument.addParagraph("The second marker is a source-bound canary.");
  sourceDocument.addImage({
    name: "sibling-image",
    dataUrl: PNG,
    alt: "Sibling image must stay unchanged",
    widthPx: 40,
    heightPx: 30,
  });
  sourceDocument.addParagraph("Review the native render before delivery.");

  const sourcePath = path.join(outputDir, "source.docx");
  await (await DocumentFile.exportDocx(sourceDocument)).save(sourcePath);
  const sourceBytes = await fs.readFile(sourcePath);
  const sourceImported = await DocumentFile.importDocx(await FileBlob.load(sourcePath));
  const sourceImages = imageProjection(sourceImported);
  const target = sourceImages.find((image) => image.alt === "Existing & approved quarterly chart");
  assert.deepEqual(target, {
    id: sourceImported.blocks[1].id,
    blockIndex: 1,
    alt: "Existing & approved quarterly chart",
    widthPx: 48,
    heightPx: 36,
    placement: {
      type: "floating",
      horizontal: { relativeTo: "margin", offsetPx: 36 },
      vertical: { relativeTo: "paragraph", offsetPx: 6 },
      wrap: "square",
      wrapSide: "bothSides",
      distanceFromTextPx: { top: 0, right: 4, bottom: 0, left: 4 },
    },
    dataUrlHash: createHash("sha256").update(sourceImported.blocks[1].dataUrl).digest("hex"),
  });
  const replacementAlt = "Q2 \"approved\" chart & variance";

  const workflowPath = path.join(repoRoot, "skills", "documents", "skills", "documents", "examples", "officekit-image-alt-text-edit-workflow.mjs");
  const { editImportedImageAltText, parseImageAltTextEditCli, imageAltTextCliOutput } = await import(workflowPath);
  const outputPath = path.join(outputDir, "output.docx");
  const auditPath = path.join(outputDir, "audit.json");
  const result = await editImportedImageAltText({
    inputPath: sourcePath,
    outputPath,
    auditPath,
    imageBlockIndex: 1,
    expectedAlt: target.alt,
    replacementAlt,
  });
  assert.equal(result.audit.provider.actual, "office-kit");
  assert.equal(result.audit.provider.silentFallback, false);
  assert.deepEqual(result.audit.savePolicy, { strategy: "rewrite", noReplace: true });
  assert.equal(result.audit.operation.type, "source-bound-image-alt-text-edit");
  assert.deepEqual(result.audit.operation.target, { id: target.id, blockIndex: 1, imageOrdinal: 0 });
  assert.equal(result.audit.operation.sourceAlt, target.alt);
  assert.equal(result.audit.operation.replacementAlt, replacementAlt);
  assert.equal(result.audit.operation.retained.widthPx, 48);
  assert.equal(result.audit.operation.retained.heightPx, 36);
  assert.deepEqual(result.audit.operation.retained.placement, target.placement);
  assert.deepEqual(result.audit.validation.changedParts, ["word/document.xml"]);
  assert.equal(result.audit.validation.imageAltTextXmlResidual.ok, true);
  assert.equal(result.audit.validation.reimport.alt, replacementAlt);
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
  const sourceDescriptions = rawImageDescriptions(sourceXml);
  const outputDescriptions = rawImageDescriptions(outputXml);
  assert.deepEqual(sourceDescriptions, [
    { docPr: target.alt, nonVisual: target.alt },
    { docPr: "Sibling image must stay unchanged", nonVisual: "Sibling image must stay unchanged" },
  ]);
  assert.deepEqual(outputDescriptions, [
    { docPr: replacementAlt, nonVisual: replacementAlt },
    { docPr: "Sibling image must stay unchanged", nonVisual: "Sibling image must stay unchanged" },
  ]);

  const reimported = await DocumentFile.importDocx(await FileBlob.load(outputPath));
  const expectedImages = structuredClone(sourceImages);
  expectedImages[0].alt = replacementAlt;
  assert.deepEqual(imageProjection(reimported), expectedImages);
  assert.equal(reimported.resolve(target.id)?.alt, replacementAlt);

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
  assert.equal(outputQa.summary.modelPixelDiff.changed, false);
  assert.equal(outputQa.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  if (nativeStatus.available) {
    assert.equal(outputQa.summary.nativeRender.pageCountMatches, true);
    assert.equal(outputQa.summary.nativeRender.pages.every((page) => page.pixelDiff.changed === false), true);
  }

  const cliOutput = path.join(outputDir, "cli-output.docx");
  const cliAudit = path.join(outputDir, "cli-audit.json");
  const cliReplacement = "Final accessible inline chart description";
  assert.deepEqual(parseImageAltTextEditCli([
    sourcePath, cliOutput, cliAudit, "3", "Sibling image must stay unchanged", cliReplacement,
  ]), {
    inputPath: sourcePath,
    outputPath: cliOutput,
    auditPath: cliAudit,
    imageBlockIndex: 3,
    expectedAlt: "Sibling image must stay unchanged",
    replacementAlt: cliReplacement,
  });
  const cli = spawnSync(process.execPath, [
    workflowPath, sourcePath, cliOutput, cliAudit, "3", "Sibling image must stay unchanged", cliReplacement,
  ], { cwd: repoRoot, encoding: "utf8" });
  assert.equal(cli.status, 0, cli.stderr);
  assert.deepEqual(JSON.parse(cli.stdout), {
    outputPath: cliOutput,
    auditPath: cliAudit,
    outputSha256: createHash("sha256").update(await fs.readFile(cliOutput)).digest("hex"),
    changedParts: ["word/document.xml"],
  });
  assert.equal((await DocumentFile.importDocx(await FileBlob.load(cliOutput))).blocks[3]?.alt, cliReplacement);

  await assert.rejects(
    () => editImportedImageAltText({
      inputPath: sourcePath,
      outputPath: path.join(outputDir, "mismatched.docx"),
      auditPath: path.join(outputDir, "mismatched.json"),
      imageBlockIndex: 1,
      expectedAlt: "Not the source description",
      replacementAlt,
    }),
    /alternative text does not match the expected source value/,
  );
  await assert.rejects(
    () => editImportedImageAltText({
      inputPath: sourcePath,
      outputPath: path.join(outputDir, "noop.docx"),
      auditPath: path.join(outputDir, "noop.json"),
      imageBlockIndex: 1,
      expectedAlt: target.alt,
      replacementAlt: target.alt,
    }),
    /replacementAlt must differ from expectedAlt/,
  );
  await assert.rejects(
    () => editImportedImageAltText({
      inputPath: sourcePath,
      outputPath: path.join(outputDir, "empty.docx"),
      auditPath: path.join(outputDir, "empty.json"),
      imageBlockIndex: 1,
      expectedAlt: target.alt,
      replacementAlt: "",
    }),
    /replacementAlt must be a non-empty alternative-text string/,
  );
  const mismatchedNativePath = path.join(outputDir, "mismatched-native.docx");
  const mismatchedNativeZip = await JSZip.loadAsync(sourceBytes);
  mismatchedNativeZip.file(
    "word/document.xml",
    sourceXml.replace(/(<pic:cNvPr\b[^>]*\bdescr=")Existing &amp; approved quarterly chart/, "$1Divergent native description"),
  );
  await fs.writeFile(mismatchedNativePath, await mismatchedNativeZip.generateAsync({ type: "nodebuffer" }), { flag: "wx" });
  await assert.rejects(
    () => editImportedImageAltText({
      inputPath: mismatchedNativePath,
      outputPath: path.join(outputDir, "mismatched-native-output.docx"),
      auditPath: path.join(outputDir, "mismatched-native-audit.json"),
      imageBlockIndex: 1,
      expectedAlt: target.alt,
      replacementAlt,
    }),
    /descriptions do not agree/,
  );
  assert.deepEqual(await fs.readFile(sourcePath), sourceBytes);
} finally {
  await fs.rm(outputDir, { recursive: true, force: true });
}

console.log("document image-alt-text workflow smoke ok");
