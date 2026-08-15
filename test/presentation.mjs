import assert from "node:assert/strict";
import path from "node:path";
import JSZip from "jszip";

import {
  column,
  DocumentFile,
  DocumentModel,
  FileBlob,
  paragraph,
  Presentation,
  PresentationFile,
  row,
  run,
  shape as composeShape,
  SpreadsheetFile,
  Workbook,
} from "../src/index.mjs";
import {
  effectivePresentationImageCrop,
  presentationImageDataUrlDimensions,
} from "../src/presentation/image-crop.mjs";
import { materializePresentationNativeGraphs } from "../src/codecs/office-kit-presentation-native.mjs";

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";
const PNG_ALT = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2nGQAAAAASUVORK5CYII=";
const JPEG = "data:image/jpeg;base64,/9j/2Q==";
const WIDE_SVG = `data:image/svg+xml;base64,${Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="400" height="200" viewBox="0 0 400 200"><rect width="200" height="200" fill="#2563eb"/><rect x="200" width="200" height="200" fill="#f97316"/></svg>').toString("base64")}`;
const TALL_SVG = `data:image/svg+xml;base64,${Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 400"><rect width="200" height="200" fill="#2563eb"/><rect y="200" width="200" height="200" fill="#f97316"/></svg>').toString("base64")}`;

// Native-object materialization must share the bounded OOXML ZIP loader. A
// selected opaque part is enough to force source-package extraction, while a
// deliberately high declared compression ratio must fail before JSZip inflates
// the source snapshot.
const boundedNativeZip = new JSZip();
boundedNativeZip.file("ppt/native.bin", "A".repeat(10_000));
const boundedNativeBytes = await boundedNativeZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
await assert.rejects(
  materializePresentationNativeGraphs({
    payload: {
      case: "presentation",
      value: { slides: [{ elements: [{ content: { case: "opaque", value: { preservedPartPaths: ["ppt/native.bin"] } } }] }] },
    },
    opaqueOpc: {
      parts: [{ path: "ppt/native.bin", contentType: "application/octet-stream" }],
      sourcePackage: { data: boundedNativeBytes },
    },
  }, { maxCompressionRatio: 1.5 }),
  /maxCompressionRatio/,
  "native source-package materialization must reject a hostile compression ratio before inflation",
);

function itemByName(items, name) {
  const item = items.find((candidate) => candidate.name === name);
  assert.ok(item, "Missing presentation object " + name);
  return item;
}

async function orderedPptxSlidePaths(zip) {
  const presentationXml = await zip.file("ppt/presentation.xml").async("text");
  const relationshipsXml = await zip.file("ppt/_rels/presentation.xml.rels").async("text");
  const relationshipTargets = new Map(
    [...relationshipsXml.matchAll(/<Relationship\b[^>]*>/g)].map(([tag]) => [
      /\bId="([^"]+)"/.exec(tag)?.[1],
      /\bTarget="([^"]+)"/.exec(tag)?.[1],
    ]),
  );
  return [...presentationXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sldId\b[^>]*\br:id="([^"]+)"[^>]*\/?\s*>/g)].map(([, relationshipId]) => {
    const target = relationshipTargets.get(relationshipId);
    assert.ok(target, `Missing SlidePart relationship ${relationshipId}`);
    return target.startsWith("/") ? target.replace(/^\/+/, "") : path.posix.normalize(path.posix.join("ppt", target));
  });
}

function relationshipPartPath(partPath) {
  const directory = path.posix.dirname(partPath);
  return path.posix.join(directory, "_rels", `${path.posix.basename(partPath)}.rels`);
}

async function assertOnlyDeclaredPptxFootprintChanged(source, output, operation) {
  const operations = Array.isArray(operation) ? operation : [operation];
  assert.equal(operations.length > 0, true);
  const partPaths = new Set(operations.map((item) => item.slidePartPath));
  assert.equal(partPaths.size, 1, "one footprint assertion must stay within one declared part");
  const sourceZip = await JSZip.loadAsync(source.bytes);
  const outputZip = await JSZip.loadAsync(output.bytes);
  const partPath = operations[0].slidePartPath;
  const sourcePart = Buffer.from(await sourceZip.file(partPath).async("uint8array"));
  const outputPart = Buffer.from(await outputZip.file(partPath).async("uint8array"));
  const masks = [];
  for (const item of operations) {
    assert.notEqual(item.leafKind, "text", "byte-offset footprint assertions currently cover scalar tokens");
    const sourceStart = Number(item.footprint.sourceStartOffset);
    const sourceEnd = Number(item.footprint.sourceEndOffset);
    const outputEnd = Number(item.footprint.outputEndOffset);
    const expected = Buffer.from(String(item.expectedValue), "utf8");
    const replacement = Buffer.from(String(item.value), "utf8");
    const outputStart = outputEnd - replacement.length;
    assert.deepEqual(sourcePart.subarray(sourceStart, sourceEnd), expected);
    assert.deepEqual(outputPart.subarray(outputStart, outputEnd), replacement);
    masks.push({ start: outputStart, end: outputEnd, bytes: expected });
  }
  let masked = outputPart;
  for (const mask of masks.sort((left, right) => right.start - left.start)) {
    masked = Buffer.concat([masked.subarray(0, mask.start), mask.bytes, masked.subarray(mask.end)]);
  }
  assert.deepEqual(masked, sourcePart, "masking all declared scalar tokens must recover the source part exactly");
  for (const [entryPath, entry] of Object.entries(sourceZip.files)) {
    if (entry.dir || entryPath === partPath) continue;
    assert.deepEqual(
      await outputZip.file(entryPath).async("uint8array"),
      await sourceZip.file(entryPath).async("uint8array"),
      `Edit Plan changed non-target part ${entryPath}`,
    );
  }
}

assert.deepEqual(presentationImageDataUrlDimensions(WIDE_SVG), { width: 400, height: 200 });
assert.deepEqual(effectivePresentationImageCrop({ fit: "cover", dataUrl: WIDE_SVG, frame: { width: 200, height: 200 } }), { left: 0.25, top: 0, right: 0.25, bottom: 0 });
assert.deepEqual(effectivePresentationImageCrop({ fit: "contain", dataUrl: WIDE_SVG, frame: { width: 200, height: 200 } }), { left: 0, top: -0.5, right: 0, bottom: -0.5 });
assert.deepEqual(effectivePresentationImageCrop({ fit: "cover", dataUrl: TALL_SVG, frame: { width: 200, height: 200 } }), { left: 0, top: 0.25, right: 0, bottom: 0.25 });
assert.deepEqual(effectivePresentationImageCrop({ fit: "contain", dataUrl: TALL_SVG, frame: { width: 200, height: 200 } }), { left: -0.5, top: 0, right: -0.5, bottom: 0 });
assert.deepEqual(effectivePresentationImageCrop({
  fit: "cover",
  crop: { left: 0.1, right: 0.1 },
  dataUrl: WIDE_SVG,
  frame: { width: 200, height: 200 },
}), { left: 0.25, top: 0, right: 0.25, bottom: 0 });
assert.throws(() => effectivePresentationImageCrop({ fit: "cover", dataUrl: "data:image/png;base64,AA==", frame: { width: 100, height: 100 } }), /intrinsic dimensions/);

const normalAutoFitDeck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const normalAutoFitShape = normalAutoFitDeck.slides.add({ name: "Normal AutoFit" }).shapes.add({
  name: "bounded-normal-autofit",
  position: { left: 80, top: 80, width: 480, height: 120 },
  text: "Keep this text within its fixed frame.",
  textBodyProperties: {
    autoFit: "shrinkText",
    normalAutoFit: { fontScale: 87.5, lineSpacingReduction: 12.5 },
  },
});
assert.deepEqual(normalAutoFitShape.text.bodyProperties.normalAutoFit, { fontScale: 87.5, lineSpacingReduction: 12.5 });
assert.throws(
  () => Presentation.create().slides.add().shapes.add({ text: "invalid", textBodyProperties: { autoFit: "resizeShape", normalAutoFit: { fontScale: 90 } } }),
  /require autoFit to be shrinkText/,
);
assert.throws(
  () => Presentation.create().slides.add().shapes.add({ text: "invalid", textBodyProperties: { autoFit: "shrinkText", normalAutoFit: { fontScale: 0.999 } } }),
  /between 1% and 100%/,
);
assert.throws(
  () => Presentation.create().slides.add().shapes.add({ text: "invalid", textBodyProperties: { autoFit: "shrinkText", normalAutoFit: { lineSpacingReduction: 1.0001 } } }),
  /at most three decimal places/,
);
const normalAutoFitSource = await PresentationFile.exportPptx(normalAutoFitDeck);
const normalAutoFitSourceZip = await JSZip.loadAsync(normalAutoFitSource.bytes);
const normalAutoFitSourceXml = await normalAutoFitSourceZip.file("ppt/slides/slide1.xml").async("text");
assert.match(normalAutoFitSourceXml, /<a:normAutofit\b[^>]*\bfontScale="87500"[^>]*\blnSpcReduction="12500"/);
const normalAutoFitImported = await PresentationFile.importPptx(normalAutoFitSource);
const importedNormalAutoFitShape = normalAutoFitImported.slides.getItem(0).shapes.getItemAt(0);
assert.deepEqual(importedNormalAutoFitShape.text.bodyProperties.normalAutoFit, { fontScale: 87.5, lineSpacingReduction: 12.5 });
importedNormalAutoFitShape.text.bodyProperties.normalAutoFit.fontScale = 82.125;
delete importedNormalAutoFitShape.text.bodyProperties.normalAutoFit.lineSpacingReduction;
const normalAutoFitEdited = await PresentationFile.exportPptx(normalAutoFitImported);
const normalAutoFitEditedZip = await JSZip.loadAsync(normalAutoFitEdited.bytes);
const normalAutoFitEditedXml = await normalAutoFitEditedZip.file("ppt/slides/slide1.xml").async("text");
assert.match(normalAutoFitEditedXml, /<a:normAutofit\b[^>]*\bfontScale="82125"/);
assert.doesNotMatch(normalAutoFitEditedXml, /\blnSpcReduction=/);
const normalAutoFitReimported = await PresentationFile.importPptx(normalAutoFitEdited);
assert.deepEqual(normalAutoFitReimported.slides.getItem(0).shapes.getItemAt(0).text.bodyProperties.normalAutoFit, { fontScale: 82.125 });

// The JavaScript layer remains the object model, Compose, inspect, resolve,
// semantic verification, and rendering surface.
const modelPresentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
assert.equal(modelPresentation.view.gridlinesVisible, false);
assert.equal(modelPresentation.view.guidesVisible, false);
assert.equal(modelPresentation.view.gridSpacingCxEmu, undefined);
assert.equal(modelPresentation.view.gridSpacingCyEmu, undefined);
assert.equal(modelPresentation.view.toProto(), undefined);
assert.equal(modelPresentation.view.showGridlines(), undefined);
assert.equal(modelPresentation.view.gridlinesVisible, true);
assert.equal(modelPresentation.view.toggleGridlines(), false);
assert.equal(modelPresentation.view.hideGridlines(), undefined);
assert.equal(modelPresentation.view.showGuides(), undefined);
assert.equal(modelPresentation.view.guidesVisible, true);
assert.deepEqual(modelPresentation.view.toProto(), { slideViewShowGuides: false, slideGuides: [] });
assert.equal(modelPresentation.view.toggleGuides(), false);
assert.equal(modelPresentation.view.hideGuides(), undefined);
assert.deepEqual(modelPresentation.toProto().viewProperties, { slideViewShowGuides: false, slideGuides: [] });
assert.deepEqual(modelPresentation.master.slideGuides, []);
assert.throws(() => modelPresentation.master.slideGuides.push({ orientation: "horizontal", position: 1 }), TypeError);
assert.throws(() => Presentation.create({ master: { slideGuides: [{ orientation: "diagonal", position: 1 }] } }), /horizontal or vertical/);
assert.throws(() => Presentation.create({ master: { slideGuides: [{ orientation: "horizontal", position: 1.5 }] } }), /signed 32-bit integer/);
const modelSlide = modelPresentation.slides.add({ name: "Compose model" });
assert.equal(modelSlide.hidden, false);
assert.deepEqual(modelSlide.visibilityCapability, { sourceBound: false, known: true, editable: true });
const crossRunReplaceShape = modelSlide.shapes.add({
  name: "cross-run-replace-guard",
  text: [{ runs: [{ text: "Alpha", style: { bold: true } }, { text: "Beta", style: { italic: true } }] }],
});
assert.throws(
  () => crossRunReplaceShape.text.replace("AlphaBeta", "Reviewed"),
  /matches across a run or paragraph boundary/i,
);
assert.equal(crossRunReplaceShape.text.value, "AlphaBeta");
assert.equal(modelSlide.hide(), modelSlide);
assert.equal(modelSlide.hidden, true);
assert.equal(modelSlide.show(), modelSlide);
assert.equal(modelSlide.setHidden(true), modelSlide);
assert.throws(() => modelSlide.setHidden("yes"), /hidden must be a boolean/i);
assert.throws(() => modelPresentation.slides.add({ hidden: 1 }), /hidden must be a boolean/i);
const composed = modelSlide.compose(
  column({ name: "compose-root", width: "fill", height: "fill", gap: 18, padding: { x: 24, y: 20 } }, [
    paragraph({ id: "compose/headline", name: "compose-headline", className: "text-slate-950 text-4xl font-bold" }, [
      "Canonical ",
      run({ textStyle: { bold: true, color: "#2563EB" } }, ["Office"]),
      " model",
    ]),
    row({ name: "compose-row", width: "fill", height: 120, gap: 16 }, [
      paragraph({ name: "compose-card-a-copy", width: "fill", height: "fill", className: "text-slate-700 text-lg" }, ["Inspect"]),
      paragraph({ name: "compose-card-b-copy", width: "fill", height: "fill", className: "text-slate-700 text-lg" }, ["Verify"]),
    ]),
    composeShape({ name: "compose-pill", width: 220, height: 48, geometry: "roundRect", fill: "#DBEAFE" }, ["Agent-ready"]),
  ]),
  { frame: { left: 80, top: 80, width: 760, height: 420 } },
);
assert.ok(composed.length >= 4);
assert.equal(modelPresentation.resolve("compose/headline").text.value, "Canonical Office model");
assert.match(modelPresentation.inspect({ kind: "deck,slide,textbox,shape", maxChars: 10_000 }).ndjson, /compose-card-b-copy/);
assert.match(modelPresentation.inspect({ kind: "textbox", target: "compose\/headline", maxChars: 4000 }).ndjson, /Canonical Office model/);
assert.equal(modelPresentation.verify().ok, true);
assert.equal(modelPresentation.validateLayout().ok, true);
const unsupportedThemePresentation = Presentation.create({ theme: { colors: { accent1: "#FF0000" } } });
unsupportedThemePresentation.slides.add().shapes.add({ text: "Theme model only" });
await assert.rejects(
  () => PresentationFile.exportPptx(unsupportedThemePresentation),
  /presentation theme customization/i,
);

// Shape alternative text is a non-visual p:cNvPr concern. It must survive the
// full model/wire/OfficeKit path without changing visible rendering, and an
// unrecognized cNvPr graph must preserve its bytes while refusing this one
// semantic mutation rather than making the whole shape unusable.
const shapeAccessibilityDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
const shapeAccessibilitySlide = shapeAccessibilityDeck.slides.add({ name: "Accessibility metadata" });
const shapeAccessibilityShape = shapeAccessibilitySlide.shapes.add({
  name: "decision-status",
  position: { left: 48, top: 72, width: 360, height: 88 },
  fill: "#DBEAFE",
  text: "Decision: controlled rollout",
  accessibility: {
    title: "Controlled rollout decision",
    description: "Status box explaining that the rollout is controlled.",
  },
});
shapeAccessibilitySlide.shapes.add({
  position: { left: 48, top: 184, width: 120, height: 48 },
  text: "Unnamed",
  accessibility: { description: "An intentionally unnamed ordinary shape." },
});
const shapeAccessibilityImage = shapeAccessibilitySlide.images.add({
  name: "decision-evidence",
  position: { left: 440, top: 72, width: 120, height: 88 },
  dataUrl: PNG,
  fit: "stretch",
  accessibility: {
    title: "Decision evidence image",
    description: "Blue evidence image supporting the controlled rollout decision.",
  },
});
const shapeAccessibilityConnector = shapeAccessibilitySlide.shapes.add({
  name: "decision-flow",
  geometry: "connector",
  from: shapeAccessibilityShape,
  to: shapeAccessibilitySlide.shapes.items[1],
  fromIdx: 3,
  toIdx: 1,
  accessibility: { title: "Decision flow", description: "Connector from the rollout decision to its review context." },
});
assert.deepEqual(shapeAccessibilityShape.accessibilityCapability, { sourceBound: false, editable: true, addable: true });
assert.deepEqual(shapeAccessibilityConnector.accessibilityCapability, { sourceBound: false, editable: true, addable: true });
assert.deepEqual(shapeAccessibilityImage.accessibilityCapability, { sourceBound: false, editable: true, addable: true });
assert.equal(shapeAccessibilityImage.alt, shapeAccessibilityImage.accessibility.description, "legacy image.alt must alias accessibility.description");
assert.match(shapeAccessibilityDeck.inspect({ kind: "shape", maxChars: 4_000 }).ndjson, /Controlled rollout decision/);
assert.match(shapeAccessibilityDeck.inspect({ kind: "image", maxChars: 4_000 }).ndjson, /Decision evidence image/);
assert.throws(() => shapeAccessibilityShape.setAccessibilityMetadata({}), /requires title, description, and\/or decorative/i);
assert.throws(() => shapeAccessibilityShape.setAccessibilityMetadata({ title: "" }), /1 through 1024 XML-safe characters/i);
assert.throws(() => shapeAccessibilityShape.setAccessibilityMetadata({ alt: "Not a cNvPr field" }), /does not support alt/i);
assert.throws(() => shapeAccessibilityImage.setAccessibilityMetadata({ description: "" }), /1 through 1024 XML-safe characters/i);
assert.throws(() => shapeAccessibilitySlide.images.add({ dataUrl: PNG, alt: "One description", accessibility: { description: "Another description" } }), /must match when both are provided/i);
assert.throws(() => shapeAccessibilityImage.replace({ name: "must-not-stick", alt: "One description", accessibility: { description: "Another description" } }), /must match when both are provided/i);
assert.equal(shapeAccessibilityImage.name, "decision-evidence", "a rejected image accessibility replacement must not partially mutate other fields");

const shapeAccessibilitySource = await PresentationFile.exportPptx(shapeAccessibilityDeck);
const shapeAccessibilitySourceZip = await JSZip.loadAsync(shapeAccessibilitySource.bytes);
const shapeAccessibilitySourceXml = await shapeAccessibilitySourceZip.file("ppt/slides/slide1.xml").async("text");
assert.match(shapeAccessibilitySourceXml, /<p:cNvPr\b(?=[^>]*\bname="decision-status")(?=[^>]*\btitle="Controlled rollout decision")(?=[^>]*\bdescr="Status box explaining that the rollout is controlled\.")[^>]*\/>/);
assert.match(shapeAccessibilitySourceXml, /<p:cNvPr\b(?=[^>]*\bname="decision-flow")(?=[^>]*\btitle="Decision flow")(?=[^>]*\bdescr="Connector from the rollout decision to its review context\.")[^>]*\/>/);
assert.match(shapeAccessibilitySourceXml, /<p:cNvPr\b(?=[^>]*\bname="decision-evidence")(?=[^>]*\btitle="Decision evidence image")(?=[^>]*\bdescr="Blue evidence image supporting the controlled rollout decision\.")[^>]*\/>/);

const shapeAccessibilityImported = await PresentationFile.importPptx(shapeAccessibilitySource);
const importedAccessibilityShape = itemByName(shapeAccessibilityImported.slides.getItem(0).shapes.items, "decision-status");
const importedAccessibilityConnector = itemByName(shapeAccessibilityImported.slides.getItem(0).connectors.items, "decision-flow");
const importedAccessibilityImage = itemByName(shapeAccessibilityImported.slides.getItem(0).images.items, "decision-evidence");
assert.deepEqual(importedAccessibilityShape.accessibility, shapeAccessibilityShape.accessibility);
assert.deepEqual(importedAccessibilityConnector.accessibility, shapeAccessibilityConnector.accessibility);
assert.deepEqual(importedAccessibilityImage.accessibility, shapeAccessibilityImage.accessibility);
assert.deepEqual(importedAccessibilityShape.accessibilityCapability, { sourceBound: true, editable: true, addable: true });
assert.deepEqual(importedAccessibilityConnector.accessibilityCapability, { sourceBound: true, editable: true, addable: true });
assert.deepEqual(importedAccessibilityImage.accessibilityCapability, { sourceBound: true, editable: true, addable: true });
assert.equal(shapeAccessibilityImported.slides.getItem(0).shapes.items[1].name, "");
assert.equal(shapeAccessibilityImported.slides.getItem(0).shapes.items[1].accessibilityCapability.editable, true);
const shapeAccessibilityNoOp = await PresentationFile.exportPptx(shapeAccessibilityImported);
assert.deepEqual(shapeAccessibilityNoOp.bytes, shapeAccessibilitySource.bytes, "unchanged imported accessibility metadata must return the exact source package");

const sourceAccessibilitySvg = await shapeAccessibilityImported.slides.getItem(0).export({ format: "svg" });
importedAccessibilityShape.setAccessibilityMetadata({ title: "Go decision: controlled rollout", description: null });
importedAccessibilityConnector.setAccessibilityMetadata({ title: null, description: "Reviewed connector from the rollout decision to context." });
importedAccessibilityImage.alt = "Reviewed evidence image for the controlled rollout decision.";
assert.equal(importedAccessibilityImage.accessibility.description, importedAccessibilityImage.alt);
importedAccessibilityImage.setAccessibilityMetadata({ title: "Reviewed decision evidence" });
const shapeAccessibilityEdited = await PresentationFile.exportPptx(shapeAccessibilityImported);
const shapeAccessibilityEditedZip = await JSZip.loadAsync(shapeAccessibilityEdited.bytes);
const shapeAccessibilityEditedXml = await shapeAccessibilityEditedZip.file("ppt/slides/slide1.xml").async("text");
assert.match(shapeAccessibilityEditedXml, /<p:cNvPr\b(?=[^>]*\bname="decision-status")(?=[^>]*\btitle="Go decision: controlled rollout")[^>]*\/>/);
assert.doesNotMatch(shapeAccessibilityEditedXml, /<p:cNvPr\b(?=[^>]*\bname="decision-status")[^>]*\bdescr=/);
assert.match(shapeAccessibilityEditedXml, /<p:cNvPr\b(?=[^>]*\bname="decision-flow")(?=[^>]*\bdescr="Reviewed connector from the rollout decision to context\.")[^>]*\/>/);
assert.doesNotMatch(shapeAccessibilityEditedXml, /<p:cNvPr\b(?=[^>]*\bname="decision-flow")[^>]*\btitle=/);
assert.match(shapeAccessibilityEditedXml, /<p:cNvPr\b(?=[^>]*\bname="decision-evidence")(?=[^>]*\btitle="Reviewed decision evidence")(?=[^>]*\bdescr="Reviewed evidence image for the controlled rollout decision\.")[^>]*\/>/);
for (const [partPath, entry] of Object.entries(shapeAccessibilitySourceZip.files)) {
  if (entry.dir || partPath === "ppt/slides/slide1.xml") continue;
  assert.deepEqual(
    await shapeAccessibilityEditedZip.file(partPath).async("uint8array"),
    await shapeAccessibilitySourceZip.file(partPath).async("uint8array"),
    `shape accessibility edit changed non-target part ${partPath}`,
  );
}
const shapeAccessibilityRoundTrip = await PresentationFile.importPptx(shapeAccessibilityEdited);
assert.deepEqual(itemByName(shapeAccessibilityRoundTrip.slides.getItem(0).shapes.items, "decision-status").accessibility, { title: "Go decision: controlled rollout" });
assert.deepEqual(itemByName(shapeAccessibilityRoundTrip.slides.getItem(0).connectors.items, "decision-flow").accessibility, { description: "Reviewed connector from the rollout decision to context." });
const roundTripAccessibilityImage = itemByName(shapeAccessibilityRoundTrip.slides.getItem(0).images.items, "decision-evidence");
assert.deepEqual(roundTripAccessibilityImage.accessibility, {
  title: "Reviewed decision evidence",
  description: "Reviewed evidence image for the controlled rollout decision.",
});
assert.equal(roundTripAccessibilityImage.alt, roundTripAccessibilityImage.accessibility.description);
const outputAccessibilitySvg = await shapeAccessibilityRoundTrip.slides.getItem(0).export({ format: "svg" });
assert.deepEqual(outputAccessibilitySvg.bytes, sourceAccessibilitySvg.bytes, "shape accessibility edits must not alter model SVG output");

const irregularShapeAccessibilityZip = await JSZip.loadAsync(shapeAccessibilitySource.bytes);
const irregularShapeAccessibilityXml = (await irregularShapeAccessibilityZip.file("ppt/slides/slide1.xml").async("text"))
  .replace(/(<p:cNvPr\b[^>]*\bname="decision-status")/, '$1 xmlns:fixture="urn:office-kit:shape-accessibility" fixture:opaque="kept"');
irregularShapeAccessibilityZip.file("ppt/slides/slide1.xml", irregularShapeAccessibilityXml);
const irregularShapeAccessibilityFile = new FileBlob(
  await irregularShapeAccessibilityZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }),
  { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
);
const irregularShapeAccessibilityImported = await PresentationFile.importPptx(irregularShapeAccessibilityFile);
const irregularAccessibilityShape = itemByName(irregularShapeAccessibilityImported.slides.getItem(0).shapes.items, "decision-status");
assert.equal(irregularAccessibilityShape.accessibility, undefined);
assert.deepEqual(irregularAccessibilityShape.accessibilityCapability, { sourceBound: true, editable: false, addable: false });
assert.throws(() => irregularAccessibilityShape.setAccessibilityMetadata({ title: "Do not rewrite unmodeled cNvPr" }), /source-bound.*editable p:cNvPr profile/i);
const irregularShapeNoOp = await PresentationFile.exportPptx(irregularShapeAccessibilityImported);
assert.deepEqual(irregularShapeNoOp.bytes, irregularShapeAccessibilityFile.bytes, "a source-bound no-op must return the exact original PPTX bytes");

const nativeLeafSourceFree = Presentation.create();
nativeLeafSourceFree.slides.add().shapes.add({ text: "Source-free" });
assert.throws(
  () => nativeLeafSourceFree.inspect({ includeNativeLeaves: true }),
  (error) => error?.code === "presentation_native_leaf_source_required",
);
assert.throws(
  () => nativeLeafSourceFree.editNativeLeaf("shape", "leaf", { expectedHash: "0".repeat(64), value: "No" }),
  (error) => error?.code === "presentation_native_leaf_source_required",
);

const nativeLeafImported = await PresentationFile.importPptx(irregularShapeAccessibilityFile);
const nativeLeafRecords = nativeLeafImported.inspect({ includeNativeLeaves: true }).ndjson
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line))
  .filter((record) => record.kind === "nativeLeaf");
const controlledTextLeaf = nativeLeafRecords.find((record) => record.targetId === irregularAccessibilityShape.id && record.value === "Decision: controlled rollout");
assert.ok(controlledTextLeaf);
assert.match(controlledTextLeaf.id, /^nl_[a-f0-9]{32}$/u);
assert.match(controlledTextLeaf.expectedHash, /^[a-f0-9]{64}$/u);
assert.match(controlledTextLeaf.revisionSha256, /^[a-f0-9]{64}$/u);
assert.equal("slidePartPath" in controlledTextLeaf, false);
assert.equal("rawXml" in controlledTextLeaf, false);
assert.equal("xpath" in controlledTextLeaf, false);
assert.throws(
  () => nativeLeafImported.editNativeLeaf(controlledTextLeaf.targetId, controlledTextLeaf.leafId, { expectedHash: "0".repeat(64), value: "Decision: hash bypass" }),
  (error) => error?.code === "presentation_native_leaf_stale",
);
assert.throws(
  () => nativeLeafImported.editNativeLeaf(`${controlledTextLeaf.targetId}-other`, controlledTextLeaf.leafId, { expectedHash: controlledTextLeaf.expectedHash, value: "Decision: cross target" }),
  (error) => error?.code === "presentation_native_leaf_not_issued",
);
assert.throws(
  () => nativeLeafImported.editNativeLeaf(controlledTextLeaf.targetId, controlledTextLeaf.leafId, { expectedHash: controlledTextLeaf.expectedHash, value: "Decision: raw bypass", rawXml: "<a:t>unsafe</a:t>" }),
  (error) => error?.code === "invalid_presentation_native_leaf_edit",
);
const controlledTextEdit = nativeLeafImported.editNativeLeaf(controlledTextLeaf.targetId, controlledTextLeaf.leafId, {
  expectedHash: controlledTextLeaf.expectedHash,
  value: "Decision: controlled native leaf",
});
assert.equal(controlledTextEdit.kind, "nativeLeafEdit");
const nativeLeafOutput = await PresentationFile.exportPptx(nativeLeafImported);
assert.equal(nativeLeafOutput.metadata.editPlan?.schema, "office-kit/pptx-edit-plan/v1");
const nativeLeafRoundTrip = await PresentationFile.importPptx(nativeLeafOutput);
assert.equal(nativeLeafRoundTrip.resolve(controlledTextLeaf.targetId).text.value, "Decision: controlled native leaf");
assert.throws(
  () => nativeLeafRoundTrip.editNativeLeaf(controlledTextLeaf.targetId, controlledTextLeaf.leafId, { expectedHash: controlledTextLeaf.expectedHash, value: "Decision: stale revision" }),
  (error) => error?.code === "presentation_native_leaf_not_issued",
);

const nativeImageGeometryImported = await PresentationFile.importPptx(irregularShapeAccessibilityFile);
const nativeImageGeometry = itemByName(nativeImageGeometryImported.slides.getItem(0).images.items, "decision-evidence");
const nativeImageGeometryLeaves = nativeImageGeometryImported.inspect({ includeNativeLeaves: true, target: nativeImageGeometry.id }).ndjson
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line))
  .filter((record) => record.kind === "nativeLeaf");
assert.deepEqual(new Set(nativeImageGeometryLeaves.map((record) => record.leafKind)), new Set(["leftEmu", "topEmu", "widthEmu", "heightEmu"]));
const nativeImageLeftLeaf = nativeImageGeometryLeaves.find((record) => record.leafKind === "leftEmu");
assert.ok(nativeImageLeftLeaf);
const nativeImageNextLeft = nativeImageLeftLeaf.value + 9_525;
nativeImageGeometryImported.editNativeLeaf(nativeImageLeftLeaf.targetId, nativeImageLeftLeaf.leafId, {
  expectedHash: nativeImageLeftLeaf.expectedHash,
  value: nativeImageNextLeft,
});
const nativeImageGeometryOutput = await PresentationFile.exportPptx(nativeImageGeometryImported);
const nativeImageGeometryOperation = nativeImageGeometryOutput.metadata.editPlan.operations[0];
assert.equal(nativeImageGeometryOperation.leafKind, "leftEmu");
await assertOnlyDeclaredPptxFootprintChanged(irregularShapeAccessibilityFile, nativeImageGeometryOutput, nativeImageGeometryOperation);
const nativeImageGeometryRoundTrip = await PresentationFile.importPptx(nativeImageGeometryOutput);
assert.equal(nativeImageGeometryRoundTrip.resolve(nativeImageGeometry.id).position.left, nativeImageNextLeft / 9_525);

const nativeShapeGeometryDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
nativeShapeGeometryDeck.slides.add().shapes.add({
  name: "multi-geometry",
  position: { left: 48, top: 72, width: 360, height: 88 },
  fill: "#DBEAFE",
  text: "Move and resize",
});
const nativeShapeGeometrySource = await PresentationFile.exportPptx(nativeShapeGeometryDeck);
const nativeShapeGeometryImported = await PresentationFile.importPptx(nativeShapeGeometrySource);
const nativeShapeGeometry = itemByName(nativeShapeGeometryImported.slides.getItem(0).shapes.items, "multi-geometry");
const nativeShapeGeometryLeaves = nativeShapeGeometryImported.inspect({ includeNativeLeaves: true, target: nativeShapeGeometry.id }).ndjson
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line))
  .filter((record) => record.kind === "nativeLeaf");
const nativeShapeLeftLeaf = nativeShapeGeometryLeaves.find((record) => record.leafKind === "leftEmu");
const nativeShapeWidthLeaf = nativeShapeGeometryLeaves.find((record) => record.leafKind === "widthEmu");
assert.ok(nativeShapeLeftLeaf);
assert.ok(nativeShapeWidthLeaf);
nativeShapeGeometryImported.editNativeLeaf(nativeShapeLeftLeaf.targetId, nativeShapeLeftLeaf.leafId, {
  expectedHash: nativeShapeLeftLeaf.expectedHash,
  value: 5,
});
nativeShapeGeometryImported.editNativeLeaf(nativeShapeWidthLeaf.targetId, nativeShapeWidthLeaf.leafId, {
  expectedHash: nativeShapeWidthLeaf.expectedHash,
  value: 123_456_789,
});
const nativeShapeGeometryOutput = await PresentationFile.exportPptx(nativeShapeGeometryImported);
const nativeShapeGeometryOperations = nativeShapeGeometryOutput.metadata.editPlan.operations;
assert.deepEqual(nativeShapeGeometryOperations.map((operation) => operation.leafKind), ["leftEmu", "widthEmu"]);
await assertOnlyDeclaredPptxFootprintChanged(nativeShapeGeometrySource, nativeShapeGeometryOutput, nativeShapeGeometryOperations);
const nativeShapeGeometryRoundTrip = await PresentationFile.importPptx(nativeShapeGeometryOutput);
assert.equal(nativeShapeGeometryRoundTrip.resolve(nativeShapeGeometry.id).position.left, 5 / 9_525);
assert.equal(nativeShapeGeometryRoundTrip.resolve(nativeShapeGeometry.id).position.width, 123_456_789 / 9_525);

const nativeChartDeck = Presentation.create({ slideSize: { width: 960, height: 540 } });
nativeChartDeck.slides.add({ name: "Native chart title" }).charts.add("bar", {
  name: "native-source-chart",
  title: "Native chart proof",
  position: { left: 120, top: 90, width: 640, height: 340 },
  categories: ["A", "B"],
  series: [{ name: "Evidence", values: [8, 13] }],
});
const nativeChartAuthored = await PresentationFile.exportPptx(nativeChartDeck);
const nativeChartSourceZip = await JSZip.loadAsync(nativeChartAuthored.bytes);
const nativeChartPartPath = Object.keys(nativeChartSourceZip.files).find((name) => /(?:^|\/)charts\/chart[0-9]+[.]xml$/iu.test(name));
assert.ok(nativeChartPartPath);
const nativeChartLiteralXml = await nativeChartSourceZip.file(nativeChartPartPath).async("text");
const nativeChartFormulaXml = nativeChartLiteralXml.replace(
  /<c:numLit>(?<body>.*?)<\/c:numLit>/su,
  "<c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache>$<body></c:numCache></c:numRef>",
);
assert.notEqual(nativeChartFormulaXml, nativeChartLiteralXml);
nativeChartSourceZip.file(nativeChartPartPath, nativeChartFormulaXml);
const nativeChartSource = new FileBlob(
  await nativeChartSourceZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }),
  { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
);
const nativeChartImported = await PresentationFile.importPptx(nativeChartSource);
const nativeChartObject = itemByName(nativeChartImported.slides.getItem(0).nativeObjects.items, "native-source-chart");
assert.equal(nativeChartObject.nativeKind, "graphicFrame");
const nativeChartLeaves = nativeChartImported.inspect({ includeNativeLeaves: true, target: nativeChartObject.id }).ndjson
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line))
  .filter((record) => record.kind === "nativeLeaf" && record.leafKind === "chartTitleText");
assert.equal(nativeChartLeaves.length, 1);
const [nativeChartTitleLeaf] = nativeChartLeaves;
assert.equal(nativeChartTitleLeaf.value, "Native chart proof");
assert.equal("targetPartPath" in nativeChartTitleLeaf, false);
assert.equal("relationshipId" in nativeChartTitleLeaf, false);
nativeChartImported.editNativeLeaf(nativeChartObject.id, nativeChartTitleLeaf.leafId, {
  expectedHash: nativeChartTitleLeaf.expectedHash,
  value: "Native chart verified",
});
const nativeChartOutput = await PresentationFile.exportPptx(nativeChartImported);
assert.deepEqual(nativeChartOutput.metadata.editPlan.changedParts, [nativeChartPartPath]);
assert.equal(nativeChartOutput.metadata.editPlan.operations.length, 1);
const [nativeChartOperation] = nativeChartOutput.metadata.editPlan.operations;
assert.equal(nativeChartOperation.leafKind, "chartTitleText");
assert.equal(nativeChartOperation.footprint.mutationPartPath, nativeChartPartPath);
const nativeChartOutputZip = await JSZip.loadAsync(nativeChartOutput.bytes);
const nativeChartOutputXml = await nativeChartOutputZip.file(nativeChartPartPath).async("text");
assert.equal(nativeChartOutputXml.replace("Native chart verified", "Native chart proof"), nativeChartFormulaXml);
for (const [partPath, entry] of Object.entries(nativeChartSourceZip.files)) {
  if (entry.dir || partPath === nativeChartPartPath) continue;
  assert.deepEqual(
    await nativeChartOutputZip.file(partPath).async("uint8array"),
    await nativeChartSourceZip.file(partPath).async("uint8array"),
    `native chart-title edit changed non-target part ${partPath}`,
  );
}
const nativeChartRoundTrip = await PresentationFile.importPptx(nativeChartOutput);
const nativeChartRoundTripObject = itemByName(nativeChartRoundTrip.slides.getItem(0).nativeObjects.items, "native-source-chart");
const nativeChartRoundTripLeaves = nativeChartRoundTrip.inspect({ includeNativeLeaves: true, target: nativeChartRoundTripObject.id }).ndjson
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line))
  .filter((record) => record.kind === "nativeLeaf" && record.leafKind === "chartTitleText");
assert.equal(nativeChartRoundTripLeaves.length, 1);
const [nativeChartRoundTripLeaf] = nativeChartRoundTripLeaves;
assert.equal(nativeChartRoundTripLeaf.value, "Native chart verified");

irregularAccessibilityShape.text.set("Decision: reviewed rollout");
const irregularOtherEdit = await PresentationFile.exportPptx(irregularShapeAccessibilityImported);
assert.equal(irregularOtherEdit.metadata.editPlan?.schema, "office-kit/pptx-edit-plan/v1");
assert.equal(irregularOtherEdit.metadata.editPlan?.operations.length, 1);
assert.deepEqual(irregularOtherEdit.metadata.editPlan?.changedParts, ["ppt/slides/slide1.xml"]);
const irregularOtherEditXml = await (await JSZip.loadAsync(irregularOtherEdit.bytes)).file("ppt/slides/slide1.xml").async("text");
assert.match(irregularOtherEditXml, /fixture:opaque="kept"/);
assert.match(irregularOtherEditXml, /title="Controlled rollout decision"/);
assert.equal(
  irregularOtherEditXml.replace("Decision: reviewed rollout", "Decision: controlled rollout"),
  irregularShapeAccessibilityXml,
  "masking the declared text leaf must recover the target SlidePart byte-for-byte",
);
const irregularOtherEditZip = await JSZip.loadAsync(irregularOtherEdit.bytes);
for (const [partPath, entry] of Object.entries(irregularShapeAccessibilityZip.files)) {
  if (entry.dir || partPath === "ppt/slides/slide1.xml") continue;
  assert.deepEqual(
    await irregularOtherEditZip.file(partPath).async("uint8array"),
    await irregularShapeAccessibilityZip.file(partPath).async("uint8array"),
    `Edit Plan changed non-target part ${partPath}`,
  );
}
irregularAccessibilityShape.accessibility = { title: "Bypass attempt" };
await assert.rejects(() => PresentationFile.exportPptx(irregularShapeAccessibilityImported), (error) => error?.code === "unsupported_presentation_edit");

// Pictures retain the older direct-alt compatibility surface. Their residual
// hash makes title/description attribute edits safe even when cNvPr carries an
// unknown child that must remain byte-owned by the source.
const irregularImageAccessibilityZip = await JSZip.loadAsync(shapeAccessibilitySource.bytes);
const irregularImageAccessibilityXml = (await irregularImageAccessibilityZip.file("ppt/slides/slide1.xml").async("text"))
  .replace(
    /(<p:cNvPr\b[^>]*\bname="decision-evidence"[^>]*)\/>/,
    '$1 xmlns:fixture="urn:office-kit:image-accessibility"><fixture:opaque value="kept"/></p:cNvPr>',
  );
irregularImageAccessibilityZip.file("ppt/slides/slide1.xml", irregularImageAccessibilityXml);
const irregularImageAccessibilityFile = new FileBlob(
  await irregularImageAccessibilityZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }),
  { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
);
const irregularImageAccessibilityImported = await PresentationFile.importPptx(irregularImageAccessibilityFile);
const irregularAccessibilityImage = itemByName(irregularImageAccessibilityImported.slides.getItem(0).images.items, "decision-evidence");
assert.deepEqual(irregularAccessibilityImage.accessibilityCapability, { sourceBound: true, editable: true, addable: true });
irregularAccessibilityImage.setAccessibilityMetadata({ title: "Reviewed opaque image metadata" });
const irregularImageAccessibilityEdited = await PresentationFile.exportPptx(irregularImageAccessibilityImported);
const irregularImageAccessibilityEditedXml = await (await JSZip.loadAsync(irregularImageAccessibilityEdited.bytes)).file("ppt/slides/slide1.xml").async("text");
assert.match(irregularImageAccessibilityEditedXml, /fixture:opaque value="kept"/);
assert.match(irregularImageAccessibilityEditedXml, /title="Reviewed opaque image metadata"/);

// Tables and charts share the same p:nvGraphicFramePr/p:cNvPr owner. Their
// alternative text is independent from visible cell/chart content and from a
// chart's visible title, while malformed native metadata remains source-owned.
const graphicFrameAccessibilityDeck = Presentation.create({ slideSize: { width: 960, height: 540 } });
const graphicFrameAccessibilitySlide = graphicFrameAccessibilityDeck.slides.add({ name: "Graphic-frame accessibility" });
const accessibleTable = graphicFrameAccessibilitySlide.tables.add({
  name: "delivery-gates",
  position: { left: 48, top: 80, width: 360, height: 160 },
  values: [["Gate", "State"], ["QA", "Pass"]],
  styleOptions: { headerRow: true },
  accessibility: {
    title: "Delivery gate table",
    description: "Two-column table listing each release gate and its current state.",
  },
});
const accessibleChart = graphicFrameAccessibilitySlide.charts.add("bar", {
  name: "regional-revenue",
  position: { left: 456, top: 80, width: 420, height: 260 },
  title: "Regional revenue",
  categories: ["North", "South"],
  series: [{ name: "Revenue", values: [42, 37] }],
  accessibility: {
    title: "Regional revenue chart",
    description: "Bar chart comparing North and South regional revenue.",
  },
});
assert.deepEqual(accessibleTable.accessibilityCapability, { sourceBound: false, editable: true, addable: true });
assert.deepEqual(accessibleChart.accessibilityCapability, { sourceBound: false, editable: true, addable: true });
assert.match(graphicFrameAccessibilityDeck.inspect({ kind: "table,chart", maxChars: 8_000 }).ndjson, /Delivery gate table/);
assert.match(graphicFrameAccessibilityDeck.inspect({ kind: "table,chart", maxChars: 8_000 }).ndjson, /Regional revenue chart/);
assert.throws(() => accessibleTable.setAccessibilityMetadata({ description: "" }), /1 through 1024 XML-safe characters/i);
assert.throws(() => accessibleChart.setAccessibilityMetadata({ decorative: true }), /cannot combine decorative: true with title or description/i);

const graphicFrameAccessibilitySource = await PresentationFile.exportPptx(graphicFrameAccessibilityDeck);
const graphicFrameAccessibilitySourceZip = await JSZip.loadAsync(graphicFrameAccessibilitySource.bytes);
const graphicFrameAccessibilitySourceXml = await graphicFrameAccessibilitySourceZip.file("ppt/slides/slide1.xml").async("text");
assert.match(graphicFrameAccessibilitySourceXml, /<p:cNvPr\b(?=[^>]*\bname="delivery-gates")(?=[^>]*\btitle="Delivery gate table")(?=[^>]*\bdescr="Two-column table listing each release gate and its current state\.")[^>]*\/>/);
assert.match(graphicFrameAccessibilitySourceXml, /<p:cNvPr\b(?=[^>]*\bname="regional-revenue")(?=[^>]*\btitle="Regional revenue chart")(?=[^>]*\bdescr="Bar chart comparing North and South regional revenue\.")[^>]*\/>/);

const graphicFrameAccessibilityImported = await PresentationFile.importPptx(graphicFrameAccessibilitySource);
const importedAccessibleTable = itemByName(graphicFrameAccessibilityImported.slides.getItem(0).tables.items, "delivery-gates");
const importedAccessibleChart = itemByName(graphicFrameAccessibilityImported.slides.getItem(0).charts.items, "regional-revenue");
assert.deepEqual(importedAccessibleTable.accessibility, accessibleTable.accessibility);
assert.deepEqual(importedAccessibleChart.accessibility, accessibleChart.accessibility);
assert.deepEqual(importedAccessibleTable.accessibilityCapability, { sourceBound: true, editable: true, addable: true });
assert.deepEqual(importedAccessibleChart.accessibilityCapability, { sourceBound: true, editable: true, addable: true });
const graphicFrameAccessibilityNoOp = await PresentationFile.exportPptx(graphicFrameAccessibilityImported);
assert.deepEqual(graphicFrameAccessibilityNoOp.bytes, graphicFrameAccessibilitySource.bytes, "unchanged graphic-frame accessibility metadata must return the exact source package");

const graphicFrameAccessibilitySourceSvg = await graphicFrameAccessibilityImported.slides.getItem(0).export({ format: "svg" });
importedAccessibleTable.setAccessibilityMetadata({ title: "Release gate table", description: null });
importedAccessibleChart.setAccessibilityMetadata({ title: null, description: "North leads South by five units." });
const graphicFrameAccessibilityEdited = await PresentationFile.exportPptx(graphicFrameAccessibilityImported);
const graphicFrameAccessibilityEditedZip = await JSZip.loadAsync(graphicFrameAccessibilityEdited.bytes);
const graphicFrameAccessibilityEditedXml = await graphicFrameAccessibilityEditedZip.file("ppt/slides/slide1.xml").async("text");
assert.match(graphicFrameAccessibilityEditedXml, /<p:cNvPr\b(?=[^>]*\bname="delivery-gates")(?=[^>]*\btitle="Release gate table")[^>]*\/>/);
assert.doesNotMatch(graphicFrameAccessibilityEditedXml, /<p:cNvPr\b(?=[^>]*\bname="delivery-gates")[^>]*\bdescr=/);
assert.match(graphicFrameAccessibilityEditedXml, /<p:cNvPr\b(?=[^>]*\bname="regional-revenue")(?=[^>]*\bdescr="North leads South by five units\.")[^>]*\/>/);
assert.doesNotMatch(graphicFrameAccessibilityEditedXml, /<p:cNvPr\b(?=[^>]*\bname="regional-revenue")[^>]*\btitle=/);
for (const [partPath, entry] of Object.entries(graphicFrameAccessibilitySourceZip.files)) {
  if (entry.dir || partPath === "ppt/slides/slide1.xml") continue;
  assert.deepEqual(
    await graphicFrameAccessibilityEditedZip.file(partPath).async("uint8array"),
    await graphicFrameAccessibilitySourceZip.file(partPath).async("uint8array"),
    `graphic-frame accessibility edit changed non-target part ${partPath}`,
  );
}
const graphicFrameAccessibilityRoundTrip = await PresentationFile.importPptx(graphicFrameAccessibilityEdited);
assert.deepEqual(itemByName(graphicFrameAccessibilityRoundTrip.slides.getItem(0).tables.items, "delivery-gates").accessibility, { title: "Release gate table" });
assert.deepEqual(itemByName(graphicFrameAccessibilityRoundTrip.slides.getItem(0).charts.items, "regional-revenue").accessibility, { description: "North leads South by five units." });
const graphicFrameAccessibilityOutputSvg = await graphicFrameAccessibilityRoundTrip.slides.getItem(0).export({ format: "svg" });
assert.deepEqual(graphicFrameAccessibilityOutputSvg.bytes, graphicFrameAccessibilitySourceSvg.bytes, "graphic-frame accessibility edits must not alter model SVG output");

const irregularGraphicFrameAccessibilityZip = await JSZip.loadAsync(graphicFrameAccessibilitySource.bytes);
const irregularGraphicFrameAccessibilityXml = (await irregularGraphicFrameAccessibilityZip.file("ppt/slides/slide1.xml").async("text"))
  .replace(/(<p:cNvPr\b[^>]*\bname="delivery-gates")/, '$1 xmlns:fixture="urn:office-kit:graphic-frame-accessibility" fixture:table="kept"')
  .replace(/(<p:cNvPr\b[^>]*\bname="regional-revenue")/, '$1 xmlns:fixture="urn:office-kit:graphic-frame-accessibility" fixture:chart="kept"');
irregularGraphicFrameAccessibilityZip.file("ppt/slides/slide1.xml", irregularGraphicFrameAccessibilityXml);
const irregularGraphicFrameAccessibilityFile = new FileBlob(
  await irregularGraphicFrameAccessibilityZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }),
  { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
);
const irregularGraphicFrameAccessibilityImported = await PresentationFile.importPptx(irregularGraphicFrameAccessibilityFile);
const irregularAccessibleTable = itemByName(irregularGraphicFrameAccessibilityImported.slides.getItem(0).tables.items, "delivery-gates");
const irregularAccessibleChart = itemByName(irregularGraphicFrameAccessibilityImported.slides.getItem(0).charts.items, "regional-revenue");
assert.equal(irregularAccessibleTable.accessibility, undefined);
assert.equal(irregularAccessibleChart.accessibility, undefined);
assert.deepEqual(irregularAccessibleTable.accessibilityCapability, { sourceBound: true, editable: false, addable: false });
assert.deepEqual(irregularAccessibleChart.accessibilityCapability, { sourceBound: true, editable: false, addable: false });
assert.throws(() => irregularAccessibleTable.setAccessibilityMetadata({ title: "Do not flatten table metadata" }), /source-bound.*editable p:cNvPr profile/i);
assert.throws(() => irregularAccessibleChart.setAccessibilityMetadata({ title: "Do not flatten chart metadata" }), /source-bound.*editable p:cNvPr profile/i);
irregularAccessibleTable.cells.set(1, 1, "Reviewed");
irregularAccessibleChart.title = "Reviewed regional revenue";
const irregularGraphicFrameOtherEdit = await PresentationFile.exportPptx(irregularGraphicFrameAccessibilityImported);
const irregularGraphicFrameOtherEditXml = await (await JSZip.loadAsync(irregularGraphicFrameOtherEdit.bytes)).file("ppt/slides/slide1.xml").async("text");
assert.match(irregularGraphicFrameOtherEditXml, /fixture:table="kept"/);
assert.match(irregularGraphicFrameOtherEditXml, /fixture:chart="kept"/);
assert.match(irregularGraphicFrameOtherEditXml, /title="Delivery gate table"/);
assert.match(irregularGraphicFrameOtherEditXml, /title="Regional revenue chart"/);
const irregularGraphicFrameBypass = await PresentationFile.importPptx(irregularGraphicFrameAccessibilityFile);
itemByName(irregularGraphicFrameBypass.slides.getItem(0).charts.items, "regional-revenue").accessibility = { title: "Bypass attempt" };
await assert.rejects(() => PresentationFile.exportPptx(irregularGraphicFrameBypass), (error) => error?.code === "unsupported_presentation_edit");

// A free-positioned line is a p:sp whose frame defines its two endpoints. It
// has no target/site identity and therefore stays separate from p:cxnSp.
const freeLineDeck = Presentation.create({ slideSize: { width: 960, height: 540 } });
const freeLineSlide = freeLineDeck.slides.add({ name: "Free line profiles" });
const horizontalFreeLine = freeLineSlide.shapes.add({
  name: "horizontal-free-line",
  geometry: "line",
  position: { left: 80, top: 90, width: 360, height: 0 },
  fill: "none",
  line: {
    style: "dash",
    fill: "#2563EB",
    width: 2,
    head: { type: "oval", width: "sm", length: "med" },
    tail: { type: "arrow", width: "lg", length: "sm" },
    cap: "round",
    join: "bevel",
  },
});
const verticalFreeLine = freeLineSlide.shapes.add({
  name: "vertical-free-line",
  geometry: "line",
  position: { left: 480, top: 80, width: 0, height: 220 },
  fill: "none",
  line: { style: "dot", fill: "#16A34A", width: 1.5 },
});
const diagonalFreeLine = freeLineSlide.shapes.add({
  name: "diagonal-free-line",
  geometry: "line",
  position: { left: 120, top: 180, width: 260, height: 160 },
  fill: "none",
  line: { style: "dashDot", fill: { color: "#F97316" }, width: 2.25 },
});
const evidenceFreeLine = freeLineSlide.shapes.add({
  name: "evidence-free-line",
  geometry: "line",
  position: { left: 520, top: 320, width: 240, height: 100 },
  fill: "none",
  line: { style: "longDashDotDot", fill: "#7C3AED", width: 2 },
});
const hiddenFreeLine = freeLineSlide.shapes.add({
  name: "hidden-free-line",
  geometry: "line",
  position: { left: 80, top: 450, width: 220, height: 0 },
  fill: "none",
  line: { style: "none", fill: "#DC2626", width: 1 },
});
assert.match(horizontalFreeLine.toSvg(), /<line\b[^>]*x1="80"[^>]*y1="90"[^>]*x2="440"[^>]*y2="90"/);
assert.match(horizontalFreeLine.toSvg(), /stroke-dasharray="8 6"/);
assert.match(horizontalFreeLine.toSvg(), /marker-start=/);
assert.match(horizontalFreeLine.toSvg(), /marker-end=/);
assert.match(horizontalFreeLine.toSvg(), /stroke-linecap="round"/);
assert.match(horizontalFreeLine.toSvg(), /stroke-linejoin="bevel"/);
assert.match(verticalFreeLine.toSvg(), /stroke-dasharray="2 4"/);
assert.match(diagonalFreeLine.toSvg(), /stroke-dasharray="8 4 2 4"/);
assert.match(diagonalFreeLine.toSvg(), /stroke="#F97316"/);
assert.match(evidenceFreeLine.toSvg(), /stroke-dasharray="8 3 2 3 2 3"/);
assert.match(hiddenFreeLine.toSvg(), /stroke="none"/);
assert.equal(freeLineSlide.connectors.items.length, 0);

const invalidFreeLineStyle = Presentation.create();
invalidFreeLineStyle.slides.add().shapes.add({
  geometry: "line",
  position: { left: 10, top: 10, width: 100, height: 0 },
  line: { style: "long-dash", fill: "#000000", width: 1 },
});
assert.throws(() => invalidFreeLineStyle.slides.getItem(0).shapes.items[0].toSvg(), /line style long-dash is unsupported/);
await assert.rejects(() => PresentationFile.exportPptx(invalidFreeLineStyle), /line style long-dash is unsupported/);
const zeroExtentFreeLine = Presentation.create();
zeroExtentFreeLine.slides.add().shapes.add({ geometry: "line", position: { left: 10, top: 10, width: 0, height: 0 } });
assert.throws(() => zeroExtentFreeLine.slides.getItem(0).shapes.items[0].toSvg(), /at least one positive extent/);
await assert.rejects(() => PresentationFile.exportPptx(zeroExtentFreeLine), /at least one positive extent/);
const placeholderFreeLine = Presentation.create();
placeholderFreeLine.slides.add().shapes.add({
  geometry: "line",
  position: { left: 10, top: 10, width: 100, height: 0 },
  placeholder: { type: "body", index: 1 },
});
await assert.rejects(() => PresentationFile.exportPptx(placeholderFreeLine), /free line.*cannot be a placeholder/i);
const nonLineArrow = Presentation.create();
nonLineArrow.slides.add().shapes.add({
  geometry: "rect",
  position: { left: 10, top: 10, width: 100, height: 40 },
  line: { fill: "#000000", width: 1, endArrow: "triangle" },
});
assert.throws(() => nonLineArrow.slides.getItem(0).shapes.items[0].toSvg(), /arrowheads require geometry line/);
await assert.rejects(() => PresentationFile.exportPptx(nonLineArrow), /arrowheads require geometry line/);
const incompleteLineEnd = Presentation.create();
incompleteLineEnd.slides.add().shapes.add({
  geometry: "line",
  position: { left: 10, top: 10, width: 100, height: 0 },
  line: { fill: "#000000", width: 1, endArrowWidth: "lg" },
});
assert.throws(() => incompleteLineEnd.slides.getItem(0).shapes.items[0].toSvg(), /requires endArrow/);
await assert.rejects(() => PresentationFile.exportPptx(incompleteLineEnd), /requires endArrow/);
const conflictingLineEnd = Presentation.create();
conflictingLineEnd.slides.add().shapes.add({
  geometry: "line",
  position: { left: 10, top: 10, width: 100, height: 0 },
  line: { fill: "#000000", width: 1, tail: "oval", endArrow: "triangle" },
});
assert.throws(() => conflictingLineEnd.slides.getItem(0).shapes.items[0].toSvg(), /conflicting tail and endArrow/);
await assert.rejects(() => PresentationFile.exportPptx(conflictingLineEnd), /conflicting tail and endArrow/);

const freeLineFirstExport = await PresentationFile.exportPptx(freeLineDeck);
const freeLineFirstZip = await JSZip.loadAsync(freeLineFirstExport.bytes);
const freeLineFirstXml = await freeLineFirstZip.file("ppt/slides/slide1.xml").async("text");
assert.equal([...freeLineFirstXml.matchAll(/<p:sp>/g)].length, 5);
assert.doesNotMatch(freeLineFirstXml, /<p:cxnSp>/);
assert.equal([...freeLineFirstXml.matchAll(/<a:prstGeom\b[^>]*\bprst="line"/g)].length, 5);
assert.match(freeLineFirstXml, /<a:ext\b[^>]*\bcx="3429000"[^>]*\bcy="0"/);
assert.match(freeLineFirstXml, /<a:ext\b[^>]*\bcx="0"[^>]*\bcy="2095500"/);
assert.match(freeLineFirstXml, /<a:ln\b[^>]*cap="rnd"/);
assert.match(freeLineFirstXml, /<a:bevel\s*\/>/);
assert.match(freeLineFirstXml, /<a:headEnd\b[^>]*type="oval"[^>]*w="sm"[^>]*len="med"/);
assert.match(freeLineFirstXml, /<a:tailEnd\b[^>]*type="arrow"[^>]*w="lg"[^>]*len="sm"/);
for (const dash of ["dash", "dot", "dashDot", "lgDashDotDot"]) {
  assert.match(freeLineFirstXml, new RegExp(`<a:prstDash val="${dash}"`));
}

const freeLineImported = await PresentationFile.importPptx(freeLineFirstExport);
const importedFreeLineSlide = freeLineImported.slides.getItem(0);
assert.equal(importedFreeLineSlide.connectors.items.length, 0);
assert.deepEqual(
  importedFreeLineSlide.shapes.items.map((shape) => [shape.name, shape.geometry, shape.line.style]),
  [
    ["horizontal-free-line", "line", "dashed"],
    ["vertical-free-line", "line", "dotted"],
    ["diagonal-free-line", "line", "dash-dot"],
    ["evidence-free-line", "line", "dash-dot-dot"],
    ["hidden-free-line", "line", "none"],
  ],
);
assert.deepEqual(itemByName(importedFreeLineSlide.shapes.items, "horizontal-free-line").position, { left: 80, top: 90, width: 360, height: 0 });
assert.deepEqual(itemByName(importedFreeLineSlide.shapes.items, "horizontal-free-line").line, {
  fill: "#2563EB",
  width: 2,
  style: "dashed",
  head: { type: "oval", width: "sm", length: "med" },
  tail: { type: "arrow", width: "lg", length: "sm" },
  cap: "round",
  join: "bevel",
});
const freeLineNoOpExport = await PresentationFile.exportPptx(freeLineImported);
const freeLineNoOpZip = await JSZip.loadAsync(freeLineNoOpExport.bytes);
assert.deepEqual(
  await freeLineNoOpZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await freeLineFirstZip.file("ppt/slides/slide1.xml").async("uint8array"),
);

const freeLineEditDeck = await PresentationFile.importPptx(freeLineFirstExport);
const freeLineToEdit = itemByName(freeLineEditDeck.slides.getItem(0).shapes.items, "horizontal-free-line");
freeLineToEdit.position.height = 48;
freeLineToEdit.line = {
  style: "dotted",
  fill: "#0F172A",
  width: 3,
  head: { type: "diamond", width: "lg", length: "lg" },
  tail: { type: "stealth", width: "sm", length: "med" },
  cap: "square",
  join: "miter",
};
const freeLineEditedRoundTrip = await PresentationFile.importPptx(await PresentationFile.exportPptx(freeLineEditDeck));
const editedFreeLine = itemByName(freeLineEditedRoundTrip.slides.getItem(0).shapes.items, "horizontal-free-line");
assert.equal(editedFreeLine.position.height, 48);
assert.deepEqual(editedFreeLine.line, {
  fill: "#0F172A",
  width: 3,
  style: "dotted",
  head: { type: "diamond", width: "lg", length: "lg" },
  tail: { type: "stealth", width: "sm", length: "med" },
  cap: "square",
  join: "miter",
});

const freeLineCloneDeck = await PresentationFile.importPptx(freeLineFirstExport);
freeLineCloneDeck.slides.getItem(0).duplicate();
const freeLineCloneRoundTrip = await PresentationFile.importPptx(await PresentationFile.exportPptx(freeLineCloneDeck));
assert.equal(freeLineCloneRoundTrip.slides.count, 2);
assert.deepEqual(
  freeLineCloneRoundTrip.slides.getItem(1).shapes.items.map((shape) => [shape.geometry, shape.line.style]),
  freeLineCloneRoundTrip.slides.getItem(0).shapes.items.map((shape) => [shape.geometry, shape.line.style]),
);

const unsupportedFreeLineZip = await JSZip.loadAsync(freeLineFirstExport.bytes);
unsupportedFreeLineZip.file("ppt/slides/slide1.xml", freeLineFirstXml.replace(/<a:prstDash val="dash"\s*\/>/, '<a:prstDash val="lgDash" />'));
const unsupportedFreeLineFile = new FileBlob(await unsupportedFreeLineZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
const unsupportedFreeLineImported = await PresentationFile.importPptx(unsupportedFreeLineFile);
const unsupportedFreeLinePreserved = await PresentationFile.exportPptx(unsupportedFreeLineImported);
const unsupportedFreeLinePreservedZip = await JSZip.loadAsync(unsupportedFreeLinePreserved.bytes);
assert.match(await unsupportedFreeLinePreservedZip.file("ppt/slides/slide1.xml").async("text"), /<a:prstDash val="lgDash"\s*\/>/);
itemByName(unsupportedFreeLineImported.slides.getItem(0).shapes.items, "horizontal-free-line").name = "Forbidden mutation";
await assert.rejects(() => PresentationFile.exportPptx(unsupportedFreeLineImported), /source-bound|read-only|unsupported/i);

// A connector endpoint is identified by both its target shape and its
// DrawingML connection-site index. The JS model keeps that pair together,
// reroutes when a modeled target moves, and preserves it through the wire and
// native package instead of silently falling back to site zero.
const connectorDeck = Presentation.create({ slideSize: { width: 960, height: 540 } });
const connectorSlide = connectorDeck.slides.add({ name: "Connector sites" });
const connectorSource = connectorSlide.shapes.add({
  name: "connector-source",
  geometry: "rect",
  position: { left: 100, top: 100, width: 200, height: 100 },
  text: "Source",
});
const connectorTarget = connectorSlide.shapes.add({
  name: "connector-target",
  geometry: "ellipse",
  position: { left: 500, top: 100, width: 200, height: 100 },
  text: "Target",
});
assert.equal(connectorSlide.shapes.getConnectionSiteIndex(connectorSource, "right"), 3);
assert.equal(connectorSlide.shapes.getConnectionSiteIndex(connectorTarget, "left"), 2);
const curvedConnector = connectorSlide.shapes.connect(connectorSource, connectorTarget, {
  name: "curved-site-connector",
  kind: "curved",
  fromSide: "right",
  toSide: "left",
  line: { style: "dashDot", fill: "#2563EB", width: 2.5 },
  head: { type: "arrow", width: "lg", length: "sm" },
  tail: { type: "diamond", width: "sm", length: "lg" },
  cap: "round",
  join: "bevel",
});
assert.deepEqual(curvedConnector.connector, {
  fromElementId: connectorSource.id,
  fromIdx: 3,
  toElementId: connectorTarget.id,
  toIdx: 2,
});
assert.deepEqual(curvedConnector.start, { x: 300, y: 150 });
assert.deepEqual(curvedConnector.end, { x: 500, y: 150 });
connectorTarget.position.left = 560;
assert.deepEqual(curvedConnector.end, { x: 560, y: 150 });
assert.equal(curvedConnector.setConnectorTo(connectorTarget, 6), curvedConnector);
assert.deepEqual(curvedConnector.end, { x: 760, y: 150 });
assert.equal(curvedConnector.setConnectorTo(connectorTarget, 2), curvedConnector);
assert.match(curvedConnector.toSvg(), / C /);
assert.match(curvedConnector.toSvg(), /stroke-dasharray="8 4 2 4"/);
assert.match(curvedConnector.toSvg(), /marker-start=/);
assert.match(curvedConnector.toSvg(), /marker-end=/);
assert.match(curvedConnector.toSvg(), /stroke-linecap="round"/);
assert.match(curvedConnector.toSvg(), /stroke-linejoin="bevel"/);
assert.throws(
  () => { curvedConnector.start = { x: 0, y: 0 }; },
  /must be changed with setConnectorFrom/,
);
curvedConnector.startSiteIndex = 4;
await assert.rejects(
  () => PresentationFile.exportPptx(connectorDeck),
  /outside the modeled rect connection-site range/,
);
curvedConnector.startSiteIndex = 3;

const hiddenConnector = connectorSlide.shapes.add({
  geometry: "connector",
  name: "hidden-site-connector",
  from: connectorSource,
  to: connectorTarget,
  fromIdx: 0,
  toIdx: 4,
  line: { style: "none", fill: "#FF0000", width: 1 },
  head: { type: "none" },
  tail: { type: "none" },
});
assert.equal(hiddenConnector.line.style, "none");
assert.match(hiddenConnector.toSvg(), /stroke="none"/);
assert.doesNotMatch(hiddenConnector.toSvg(), /marker-(?:start|end)=/);
assert.equal(curvedConnector.isForeground, false);
assert.equal(curvedConnector.bringToFront(), curvedConnector);
assert.equal(curvedConnector.isForeground, true);
assert.equal(connectorSlide.connectors.items.at(-1), curvedConnector);
assert.equal(curvedConnector.sendToBack(), curvedConnector);
assert.equal(connectorSlide.connectors.items[0], curvedConnector);

const connectorBoundaryDeck = Presentation.create();
const connectorBoundarySlide = connectorBoundaryDeck.slides.add();
const boundaryRect = connectorBoundarySlide.shapes.add({ geometry: "rect", position: { left: 10, top: 10, width: 100, height: 50 }, text: "A" });
const boundaryEllipse = connectorBoundarySlide.shapes.add({ geometry: "ellipse", position: { left: 200, top: 10, width: 100, height: 50 }, text: "B" });
assert.throws(
  () => connectorBoundarySlide.shapes.add({ geometry: "connector", from: boundaryRect, to: boundaryEllipse }),
  /requires fromIdx and toIdx/,
);
assert.throws(
  () => connectorBoundarySlide.shapes.add({ geometry: "connector", from: boundaryRect, to: boundaryEllipse, fromIdx: 4, toIdx: 2 }),
  /outside the modeled rect connection-site range/,
);
const unsupportedSiteShape = connectorBoundarySlide.shapes.add({ geometry: "custom", position: { left: 10, top: 100, width: 100, height: 50 }, text: "Custom" });
assert.throws(
  () => connectorBoundarySlide.shapes.getConnectionSiteIndex(unsupportedSiteShape, "right"),
  /no modeled connection-site map/,
);
const connectorGroup = connectorBoundarySlide.groups.add({
  name: "connector-group",
  position: { left: 300, top: 100, width: 300, height: 150 },
  childFrame: { left: 0, top: 0, width: 300, height: 150 },
});
const groupFrom = connectorGroup.shapes.add({ geometry: "rect", position: { left: 0, top: 0, width: 80, height: 50 }, text: "From" });
const groupTo = connectorGroup.shapes.add({ geometry: "roundRect", position: { left: 180, top: 0, width: 80, height: 50 }, text: "To" });
assert.throws(
  () => connectorBoundarySlide.shapes.connect(boundaryRect, groupFrom),
  /same slide or group shape tree/,
);
const groupConnector = connectorGroup.shapes.connect(groupFrom, groupTo, { fromSide: "right", toSide: "left" });
assert.equal(connectorGroup.children[0], groupConnector);
assert.equal(groupConnector.bringToFront(), groupConnector);
assert.equal(connectorGroup.children.at(-1), groupConnector);
assert.equal(groupConnector.sendToBack(), groupConnector);
assert.equal(connectorGroup.children[0], groupConnector);

const connectorFirstExport = await PresentationFile.exportPptx(connectorDeck);
const connectorFirstZip = await JSZip.loadAsync(connectorFirstExport.bytes);
const connectorFirstXml = await connectorFirstZip.file("ppt/slides/slide1.xml").async("text");
assert.ok(connectorFirstXml.indexOf("<p:cxnSp>") < connectorFirstXml.indexOf("<p:sp>"));
assert.match(connectorFirstXml, /<a:stCxn\b[^>]*idx="3"/);
assert.match(connectorFirstXml, /<a:endCxn\b[^>]*idx="2"/);
assert.match(connectorFirstXml, /prst="curvedConnector3"/);
assert.match(connectorFirstXml, /<a:prstDash val="dashDot"/);
assert.match(connectorFirstXml, /<a:ln\b[^>]*cap="rnd"/);
assert.match(connectorFirstXml, /<a:bevel\s*\/>/);
assert.match(connectorFirstXml, /<a:headEnd\b[^>]*type="arrow"[^>]*w="lg"[^>]*len="sm"/);
assert.match(connectorFirstXml, /<a:tailEnd\b[^>]*type="diamond"[^>]*w="sm"[^>]*len="lg"/);
assert.match(connectorFirstXml, /<a:noFill\s*\/>/);

const connectorImported = await PresentationFile.importPptx(connectorFirstExport);
const importedConnectorSlide = connectorImported.slides.getItem(0);
const importedCurvedConnector = itemByName(importedConnectorSlide.connectors.items, "curved-site-connector");
assert.equal(importedCurvedConnector.connectorType, "curved");
assert.equal(importedCurvedConnector.startSiteIndex, 3);
assert.equal(importedCurvedConnector.endSiteIndex, 2);
assert.deepEqual(importedCurvedConnector.head, { type: "arrow", width: "lg", length: "sm" });
assert.deepEqual(importedCurvedConnector.tail, { type: "diamond", width: "sm", length: "lg" });
assert.equal(importedCurvedConnector.line.style, "dash-dot");
assert.equal(importedCurvedConnector.cap, "round");
assert.equal(importedCurvedConnector.join, "bevel");
assert.equal(itemByName(importedConnectorSlide.connectors.items, "hidden-site-connector").line.style, "none");
assert.throws(() => importedCurvedConnector.bringToFront(), /z-order is source-bound/);
const connectorNoOpExport = await PresentationFile.exportPptx(connectorImported);
const connectorNoOpZip = await JSZip.loadAsync(connectorNoOpExport.bytes);
assert.deepEqual(
  await connectorNoOpZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await connectorFirstZip.file("ppt/slides/slide1.xml").async("uint8array"),
);

const connectorCloneDeck = await PresentationFile.importPptx(connectorFirstExport);
const connectorCloneSlide = connectorCloneDeck.slides.getItem(0).duplicate();
const clonedCurvedConnector = itemByName(connectorCloneSlide.connectors.items, "curved-site-connector");
assert.equal(clonedCurvedConnector.startSiteIndex, 3);
assert.equal(clonedCurvedConnector.endSiteIndex, 2);
const connectorCloneExport = await PresentationFile.exportPptx(connectorCloneDeck);
const connectorCloneRoundTrip = await PresentationFile.importPptx(connectorCloneExport);
const roundTripClonedConnector = itemByName(connectorCloneRoundTrip.slides.getItem(1).connectors.items, "curved-site-connector");
assert.equal(roundTripClonedConnector.startSiteIndex, 3);
assert.equal(roundTripClonedConnector.endSiteIndex, 2);

const connectorEditDeck = await PresentationFile.importPptx(connectorFirstExport);
const connectorEditSlide = connectorEditDeck.slides.getItem(0);
const connectorEditTarget = itemByName(connectorEditSlide.shapes.items, "connector-target");
const connectorToEdit = itemByName(connectorEditSlide.connectors.items, "curved-site-connector");
connectorEditTarget.position.left += 40;
connectorToEdit.setConnectorTo(connectorEditTarget, 6);
const connectorEditedExport = await PresentationFile.exportPptx(connectorEditDeck);
const connectorEditedRoundTrip = await PresentationFile.importPptx(connectorEditedExport);
const editedConnector = itemByName(connectorEditedRoundTrip.slides.getItem(0).connectors.items, "curved-site-connector");
assert.equal(editedConnector.endSiteIndex, 6);
assert.equal(editedConnector.end.x, 800);

// Custom shows are a real inline PresentationML graph. Source-free decks own
// the complete list; canonical imports may edit only names and ordered slide
// membership while show topology/native identity remain source-bound.
const customShowDeck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const customShowOverview = customShowDeck.slides.add({ name: "Overview" });
customShowOverview.shapes.add({
  name: "overview-title",
  position: { left: 80, top: 80, width: 800, height: 80 },
  text: [{ runs: [{ text: "Overview", link: { customShow: "Board route", returnToSlide: true, tooltip: "Open board route" } }] }],
});
const customShowEvidence = customShowDeck.slides.add({ name: "Evidence" });
customShowEvidence.shapes.add({ name: "evidence-title", position: { left: 80, top: 80, width: 800, height: 80 }, text: "Evidence" });
const customShowAppendix = customShowDeck.slides.add({ name: "Appendix" });
customShowAppendix.shapes.add({ name: "appendix-title", position: { left: 80, top: 80, width: 800, height: 80 }, text: "Appendix" });
const boardShow = customShowDeck.customShows.add({
  id: "custom-show/board",
  name: "Board route",
  nativeId: 7,
  slides: [customShowOverview, customShowAppendix],
});
customShowDeck.customShows.add({
  id: "custom-show/review",
  name: "Review route",
  nativeId: 11,
  slides: [customShowEvidence],
});
assert.equal(customShowDeck.resolve(boardShow.id), boardShow);
assert.match(customShowDeck.inspect({ kind: "customShow", maxChars: 4000 }).ndjson, /Board route/);
const customShowFirstExport = await PresentationFile.exportPptx(customShowDeck);
const customShowFirstZip = await JSZip.loadAsync(customShowFirstExport.bytes);
const customShowPresentationXml = await customShowFirstZip.file("ppt/presentation.xml").async("string");
assert.match(customShowPresentationXml, /<p:custShowLst>/);
assert.match(customShowPresentationXml, /<p:custShow name="Board route" id="7"><p:sldLst><p:sld r:id="rIdSlide1"[^>]*\/><p:sld r:id="rIdSlide3"[^>]*\/><\/p:sldLst><\/p:custShow>/);
assert.ok(customShowPresentationXml.indexOf("<p:custShowLst>") < customShowPresentationXml.indexOf("<p:defaultTextStyle"));
const customShowSlideXml = await customShowFirstZip.file("ppt/slides/slide1.xml").async("string");
assert.match(customShowSlideXml, /<a:hlinkClick r:id=""[^>]*action="ppaction:\/\/customshow\?id=7&amp;return=true"[^>]*tooltip="Open board route"/);

const customShowImported = await PresentationFile.importPptx(customShowFirstExport);
assert.equal(customShowImported.customShows.count, 2);
assert.deepEqual(customShowImported.customShows.getItem("Board route").slideIds, [customShowImported.slides.items[0].id, customShowImported.slides.items[2].id]);
assert.deepEqual(itemByName(customShowImported.slides.items[0].shapes.items, "overview-title").text.paragraphs[0].runs[0].link, {
  customShow: "Board route",
  returnToSlide: true,
  tooltip: "Open board route",
});
const editableBoardShow = customShowImported.customShows.getItem("Board route");
editableBoardShow.name = "Executive route";
editableBoardShow.setSlides([customShowImported.slides.items[2], customShowImported.slides.items[0], customShowImported.slides.items[2]]);
const customShowEditedExport = await PresentationFile.exportPptx(customShowImported);
const customShowEditedRoundTrip = await PresentationFile.importPptx(customShowEditedExport);
assert.equal(customShowEditedRoundTrip.customShows.count, 2);
assert.deepEqual(customShowEditedRoundTrip.customShows.getItem("Executive route").slideIds, [
  customShowEditedRoundTrip.slides.items[2].id,
  customShowEditedRoundTrip.slides.items[0].id,
  customShowEditedRoundTrip.slides.items[2].id,
]);
assert.equal(
  itemByName(customShowEditedRoundTrip.slides.items[0].shapes.items, "overview-title").text.paragraphs[0].runs[0].link.customShow,
  "Executive route",
  "renaming a custom show must retain the native link identity and refresh its public display name",
);
const customShowEditedZip = await JSZip.loadAsync(customShowEditedExport.bytes);
for (const partPath of Object.keys(customShowFirstZip.files).filter((entry) => !customShowFirstZip.files[entry].dir && entry !== "ppt/presentation.xml")) {
  assert.deepEqual(
    await customShowEditedZip.file(partPath).async("uint8array"),
    await customShowFirstZip.file(partPath).async("uint8array"),
    `custom-show edit changed non-presentation part ${partPath}`,
  );
}
const customShowRetargeted = await PresentationFile.importPptx(customShowFirstExport);
const customShowRetargetedShape = itemByName(customShowRetargeted.slides.items[0].shapes.items, "overview-title");
const customShowRetargetedParagraph = customShowRetargetedShape.text.paragraphs[0];
customShowRetargetedParagraph.runs[0].link = {
  customShow: "Review route",
  returnToSlide: false,
};
customShowRetargetedShape.text.paragraphs = [customShowRetargetedParagraph];
const customShowRetargetedExport = await PresentationFile.exportPptx(customShowRetargeted);
const customShowRetargetedZip = await JSZip.loadAsync(customShowRetargetedExport.bytes);
assert.match(await customShowRetargetedZip.file("ppt/slides/slide1.xml").async("string"), /action="ppaction:\/\/customshow\?id=11&amp;return=false"/);
const customShowRetargetedRoundTrip = await PresentationFile.importPptx(customShowRetargetedExport);
assert.deepEqual(itemByName(customShowRetargetedRoundTrip.slides.items[0].shapes.items, "overview-title").text.paragraphs[0].runs[0].link, {
  customShow: "Review route",
  returnToSlide: false,
});
const missingCustomShowLink = Presentation.create({ slideSize: { width: 640, height: 360 } });
missingCustomShowLink.slides.add().shapes.add({
  position: { left: 40, top: 40, width: 360, height: 80 },
  text: [{ runs: [{ text: "Missing", link: { customShow: "Not present" } }] }],
});
await assert.rejects(
  () => PresentationFile.exportPptx(missingCustomShowLink),
  (error) => error?.code === "invalid_presentation_hyperlink",
);
const customShowCloneImport = await PresentationFile.importPptx(customShowFirstExport);
customShowCloneImport.slides.items[0].duplicate();
customShowCloneImport.customShows.getItem("Board route").name = "Cloned executive route";
const customShowCloneExport = await PresentationFile.exportPptx(customShowCloneImport);
const customShowCloneZip = await JSZip.loadAsync(customShowCloneExport.bytes);
const customShowCloneSlidePath = (await orderedPptxSlidePaths(customShowCloneZip))[1];
assert.deepEqual(
  await customShowCloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await customShowFirstZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning a custom-show run link must preserve the retained source SlidePart byte-for-byte",
);
assert.deepEqual(
  await customShowCloneZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  await customShowFirstZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  "a relationship-free custom-show action must not rewrite the retained source relationship graph",
);
assert.match(
  await customShowCloneZip.file(customShowCloneSlidePath).async("string"),
  /<a:hlinkClick\b[^>]*r:id=""[^>]*action="ppaction:\/\/customshow\?id=7&amp;return=true"[^>]*tooltip="Open board route"/,
);
assert.doesNotMatch(
  await customShowCloneZip.file(relationshipPartPath(customShowCloneSlidePath)).async("string"),
  /relationships\/(?:hyperlink|slide)"/,
  "custom-show run links must remain relationship-free on the cloned SlidePart",
);
const customShowCloneRoundTrip = await PresentationFile.importPptx(customShowCloneExport);
assert.deepEqual(customShowCloneRoundTrip.slides.items.map((slide) => slide.name), ["Overview", "Overview", "Evidence", "Appendix"]);
assert.deepEqual(customShowCloneRoundTrip.customShows.getItem("Cloned executive route").slideIds, [
  customShowCloneRoundTrip.slides.items[0].id,
  customShowCloneRoundTrip.slides.items[3].id,
]);
assert.ok(!customShowCloneRoundTrip.customShows.getItem("Cloned executive route").slideIds.includes(customShowCloneRoundTrip.slides.items[1].id));
for (const slideIndex of [0, 1]) {
  assert.deepEqual(
    itemByName(customShowCloneRoundTrip.slides.items[slideIndex].shapes.items, "overview-title").text.paragraphs[0].runs[0].link,
    { customShow: "Cloned executive route", returnToSlide: true, tooltip: "Open board route" },
  );
}
const customShowCompoundCloneImport = await PresentationFile.importPptx(customShowFirstExport);
const customShowCompoundClone = customShowCompoundCloneImport.slides.items[0].duplicate();
customShowCompoundCloneImport.customShows.getItem("Board route").setSlides([customShowCompoundClone]);
await assert.rejects(
  () => PresentationFile.exportPptx(customShowCompoundCloneImport),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "custom-show membership changes must cross a separate export/reimport boundary from slide cloning",
);
const customShowMovedImport = await PresentationFile.importPptx(customShowFirstExport);
customShowMovedImport.slides.items[2].moveTo(0);
const customShowMovedRoundTrip = await PresentationFile.importPptx(await PresentationFile.exportPptx(customShowMovedImport));
assert.deepEqual(customShowMovedRoundTrip.slides.items.map((slide) => slide.name), ["Appendix", "Overview", "Evidence"]);
assert.deepEqual(customShowMovedRoundTrip.customShows.getItem("Board route").slides.map((slide) => slide.name), ["Overview", "Appendix"]);

const customShowAddedImport = await PresentationFile.importPptx(customShowFirstExport);
customShowAddedImport.customShows.add("Added route", [customShowAddedImport.slides.items[0]]);
await assert.rejects(
  () => PresentationFile.exportPptx(customShowAddedImport),
  (error) => error?.code === "presentation_custom_show_topology_changed",
);
const customShowIdentityImport = await PresentationFile.importPptx(customShowFirstExport);
customShowIdentityImport.customShows.items[0].nativeId = 99;
await assert.rejects(
  () => PresentationFile.exportPptx(customShowIdentityImport),
  (error) => error?.code === "presentation_custom_show_topology_changed",
);

const irregularCustomShowZip = await JSZip.loadAsync(customShowFirstExport.bytes);
const irregularCustomShowXml = (await irregularCustomShowZip.file("ppt/presentation.xml").async("string"))
  .replace("</p:custShow>", "<p:extLst/></p:custShow>");
irregularCustomShowZip.file("ppt/presentation.xml", irregularCustomShowXml);
const irregularCustomShowBytes = await irregularCustomShowZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const irregularCustomShowFile = new FileBlob(irregularCustomShowBytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
const irregularCustomShowImport = await PresentationFile.importPptx(irregularCustomShowFile);
assert.equal(irregularCustomShowImport.customShows.count, 0);
assert.equal(itemByName(irregularCustomShowImport.slides.items[0].shapes.items, "overview-title").text.paragraphs[0].runs[0].link, undefined);
const irregularLinkMutation = await PresentationFile.importPptx(irregularCustomShowFile);
const irregularLinkShape = itemByName(irregularLinkMutation.slides.items[0].shapes.items, "overview-title");
const irregularLinkParagraph = irregularLinkShape.text.paragraphs[0];
irregularLinkParagraph.runs[0].link = { uri: "https://example.com/replacement" };
irregularLinkShape.text.paragraphs = [irregularLinkParagraph];
await assert.rejects(
  () => PresentationFile.exportPptx(irregularLinkMutation),
  (error) => error?.code === "unsupported_presentation_edit",
);
irregularCustomShowImport.customShows.add("Unsafe replacement", [irregularCustomShowImport.slides.items[0]]);
await assert.rejects(
  () => PresentationFile.exportPptx(irregularCustomShowImport),
  (error) => error?.code === "unsupported_presentation_custom_show_edit",
);

// PowerPoint sections are a p14 extension graph, not a custom-show route:
// they partition the complete ordered slide sequence through native p:sldId
// values and keep their GUID identity fixed after an import.
const sectionDeck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const sectionOpening = sectionDeck.slides.add({ name: "Opening" });
sectionOpening.shapes.add({ name: "section-opening", position: { left: 80, top: 80, width: 800, height: 80 }, text: "Opening" });
const sectionEvidence = sectionDeck.slides.add({ name: "Evidence" });
sectionEvidence.shapes.add({ name: "section-evidence", position: { left: 80, top: 80, width: 800, height: 80 }, text: "Evidence" });
const sectionAppendix = sectionDeck.slides.add({ name: "Appendix" });
sectionAppendix.shapes.add({ name: "section-appendix", position: { left: 80, top: 80, width: 800, height: 80 }, text: "Appendix" });
const openingSection = sectionDeck.sections.add({
  id: "section/opening",
  name: "Opening",
  nativeId: "{01F07B81-39E6-4BBB-9B89-66EA253FBD29}",
  slides: [sectionOpening],
});
sectionDeck.sections.add({
  id: "section/content",
  name: "Content",
  nativeId: "{1FEF2C88-0CF2-4176-BA81-0DE6FD9D1274}",
  slides: [sectionEvidence, sectionAppendix],
});
assert.equal(sectionDeck.resolve(openingSection.id), openingSection);
assert.match(sectionDeck.inspect({ kind: "section", maxChars: 4000 }).ndjson, /Opening/);
assert.equal(sectionDeck.verify().ok, true);
const sectionFirstExport = await PresentationFile.exportPptx(sectionDeck);
const sectionFirstZip = await JSZip.loadAsync(sectionFirstExport.bytes);
const sectionPresentationXml = await sectionFirstZip.file("ppt/presentation.xml").async("string");
assert.match(sectionPresentationXml, /<p:ext uri="\{521415D9-36F7-43E2-AB2F-B90AF26B5E84\}"><p14:sectionLst/);
assert.match(sectionPresentationXml, /<p14:section name="Opening" id="\{01F07B81-39E6-4BBB-9B89-66EA253FBD29\}"><p14:sldIdLst><p14:sldId id="256" \/><\/p14:sldIdLst><\/p14:section>/);
assert.match(sectionPresentationXml, /<p14:section name="Content" id="\{1FEF2C88-0CF2-4176-BA81-0DE6FD9D1274\}"><p14:sldIdLst><p14:sldId id="257" \/><p14:sldId id="258" \/><\/p14:sldIdLst><\/p14:section>/);
const sectionImported = await PresentationFile.importPptx(sectionFirstExport);
assert.equal(sectionImported.sections.count, 2);
assert.deepEqual(sectionImported.sections.getItem("Opening").slideIds, [sectionImported.slides.items[0].id]);
assert.deepEqual(sectionImported.sections.getItem("Content").slideIds, [sectionImported.slides.items[1].id, sectionImported.slides.items[2].id]);
const editableOpeningSection = sectionImported.sections.getItem("Opening");
editableOpeningSection.name = "Introduction";
editableOpeningSection.setSlides([sectionImported.slides.items[0], sectionImported.slides.items[1]]);
sectionImported.sections.getItem("Content").setSlides([sectionImported.slides.items[2]]);
const sectionEditedExport = await PresentationFile.exportPptx(sectionImported);
const sectionEditedRoundTrip = await PresentationFile.importPptx(sectionEditedExport);
assert.equal(sectionEditedRoundTrip.sections.getItem("Introduction").nativeId, "{01F07B81-39E6-4BBB-9B89-66EA253FBD29}");
assert.deepEqual(sectionEditedRoundTrip.sections.getItem("Introduction").slides.map((slide) => slide.name), ["Opening", "Evidence"]);
assert.deepEqual(sectionEditedRoundTrip.sections.getItem("Content").slides.map((slide) => slide.name), ["Appendix"]);
const sectionEditedZip = await JSZip.loadAsync(sectionEditedExport.bytes);
for (const partPath of Object.keys(sectionFirstZip.files).filter((entry) => !sectionFirstZip.files[entry].dir && entry !== "ppt/presentation.xml")) {
  assert.deepEqual(
    await sectionEditedZip.file(partPath).async("uint8array"),
    await sectionFirstZip.file(partPath).async("uint8array"),
    `section edit changed non-presentation part ${partPath}`,
  );
}
const sectionAddedImport = await PresentationFile.importPptx(sectionFirstExport);
sectionAddedImport.sections.add("Unsafe", [sectionAddedImport.slides.items[0], sectionAddedImport.slides.items[1], sectionAddedImport.slides.items[2]]);
await assert.rejects(
  () => PresentationFile.exportPptx(sectionAddedImport),
  (error) => error?.code === "presentation_section_topology_changed",
);
const sectionCloneImport = await PresentationFile.importPptx(sectionFirstExport);
assert.equal(sectionCloneImport.slides.items[0].cloneCapability.supported, false);
assert.throws(
  () => sectionCloneImport.slides.items[0].duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);
assert.equal(sectionCloneImport.slides.items.length, 3);
const irregularSectionZip = await JSZip.loadAsync(sectionFirstExport.bytes);
const irregularSectionXml = (await irregularSectionZip.file("ppt/presentation.xml").async("string"))
  .replace("</p14:section>", "<p14:extLst/></p14:section>");
irregularSectionZip.file("ppt/presentation.xml", irregularSectionXml);
const irregularSectionBytes = await irregularSectionZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const irregularSectionFile = new FileBlob(irregularSectionBytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
const irregularSectionImport = await PresentationFile.importPptx(irregularSectionFile);
assert.equal(irregularSectionImport.sections.count, 0);
irregularSectionImport.sections.add("Unsafe", irregularSectionImport.slides.items);
await assert.rejects(
  () => PresentationFile.exportPptx(irregularSectionImport),
  (error) => error?.code === "unsupported_presentation_section_edit",
);

// Slide visibility is one source-bound p:sld/@show leaf. The public state is
// intentionally inverted to `hidden`, so Agent code never needs to remember
// that native show="0" means hidden and absence means visible.
const visibilityDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
visibilityDeck.slides.add({ name: "Visible slide" }).shapes.add({ text: "Visible" });
visibilityDeck.slides.add({ name: "Hidden slide", hidden: true }).shapes.add({ text: "Hidden" });
const visibilitySource = await PresentationFile.exportPptx(visibilityDeck);
const visibilitySourceZip = await JSZip.loadAsync(visibilitySource.bytes);
assert.doesNotMatch(await visibilitySourceZip.file("ppt/slides/slide1.xml").async("string"), /<p:sld\b[^>]*\bshow=/);
assert.match(await visibilitySourceZip.file("ppt/slides/slide2.xml").async("string"), /<p:sld\b[^>]*\bshow="0"/);
const visibilityImported = await PresentationFile.importPptx(visibilitySource);
assert.deepEqual(visibilityImported.slides.items.map((slide) => slide.hidden), [false, true]);
assert.deepEqual(visibilityImported.slides.items[0].visibilityCapability, { sourceBound: true, known: true, editable: true });
assert.match(visibilityImported.inspect({ kind: "slide", maxChars: 4000 }).ndjson, /"hidden":true/);
visibilityImported.slides.items[0].hide();
const visibilityEdited = await PresentationFile.exportPptx(visibilityImported);
const visibilityEditedZip = await JSZip.loadAsync(visibilityEdited.bytes);
assert.match(await visibilityEditedZip.file("ppt/slides/slide1.xml").async("string"), /<p:sld\b[^>]*\bshow="0"/);
for (const [partPath, entry] of Object.entries(visibilitySourceZip.files)) {
  if (entry.dir || partPath === "ppt/slides/slide1.xml") continue;
  assert.deepEqual(
    await visibilityEditedZip.file(partPath).async("uint8array"),
    await visibilitySourceZip.file(partPath).async("uint8array"),
    `slide visibility edit changed non-target part ${partPath}`,
  );
}
const visibilityRoundTrip = await PresentationFile.importPptx(visibilityEdited);
assert.deepEqual(visibilityRoundTrip.slides.items.map((slide) => slide.hidden), [true, true]);
visibilityRoundTrip.slides.items[0].show();
const visibilityShown = await PresentationFile.exportPptx(visibilityRoundTrip);
const visibilityShownZip = await JSZip.loadAsync(visibilityShown.bytes);
assert.doesNotMatch(await visibilityShownZip.file("ppt/slides/slide1.xml").async("string"), /<p:sld\b[^>]*\bshow=/);

const opaqueVisibilityZip = await JSZip.loadAsync(visibilitySource.bytes);
opaqueVisibilityZip.file(
  "ppt/slides/slide2.xml",
  (await opaqueVisibilityZip.file("ppt/slides/slide2.xml").async("string")).replace('show="0"', 'show="sometimes"'),
);
const opaqueVisibilityBytes = await opaqueVisibilityZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const opaqueVisibilityFile = new FileBlob(opaqueVisibilityBytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
const opaqueVisibilityImported = await PresentationFile.importPptx(opaqueVisibilityFile);
assert.deepEqual(opaqueVisibilityImported.slides.items[1].visibilityCapability, { sourceBound: true, known: false, editable: false });
assert.throws(() => opaqueVisibilityImported.slides.items[1].show(), /source-bound and not safely editable/i);
const opaqueVisibilityPreserved = await PresentationFile.exportPptx(opaqueVisibilityImported);
assert.deepEqual(opaqueVisibilityPreserved.bytes, opaqueVisibilityBytes);

// Slide transitions are a direct p:transition leaf, deliberately distinct
// from animation/timing graphs. The profile owns the complete ECMA-376 base
// effect vocabulary, explicit speed/click behavior, and an optional timer.
const transitionDeck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const transitionFade = transitionDeck.slides.add({ name: "Fade" });
transitionFade.shapes.add({ name: "transition-fade-title", position: { left: 80, top: 80, width: 800, height: 80 }, text: "Fade" });
transitionFade.setTransition({ effect: "fade", speed: "medium", durationMs: 750, advanceOnClick: true, advanceAfterMs: 1250 });
const transitionPush = transitionDeck.slides.add({ name: "Push" });
transitionPush.shapes.add({ name: "transition-push-title", position: { left: 80, top: 80, width: 800, height: 80 }, text: "Push" });
transitionPush.setTransition({ effect: "push", direction: "right", speed: "fast", advanceOnClick: false, advanceAfterMs: 0 });
const transitionWipe = transitionDeck.slides.add({ name: "Wipe" });
transitionWipe.shapes.add({ name: "transition-wipe-title", position: { left: 80, top: 80, width: 800, height: 80 }, text: "Wipe" });
transitionWipe.setTransition({ effect: "wipe", speed: "slow", advanceOnClick: true });
assert.equal(transitionDeck.resolve(`${transitionFade.id}/transition`), transitionFade.transition);
assert.deepEqual(transitionFade.transition.toJSON(), { effect: "fade", speed: "medium", durationMs: 750, advanceOnClick: true, advanceAfterMs: 1250 });
assert.match(transitionDeck.inspect({ kind: "transition", maxChars: 4000 }).ndjson, /"effect":"push"/);
assert.throws(() => transitionFade.setTransition({ effect: "fade", direction: "left" }), /does not accept direction/);
assert.throws(() => transitionPush.setTransition({ effect: "push", direction: "diagonal" }), /left, up, right, or down/);
assert.throws(() => transitionWipe.setTransition({ effect: "wipe", direction: "diagonal" }), /left, up, right, or down/);
assert.throws(() => transitionPush.setTransition({ effect: "split", direction: "diagonal" }), /in or out/);
const transitionFirstExport = await PresentationFile.exportPptx(transitionDeck);
const transitionFirstZip = await JSZip.loadAsync(transitionFirstExport.bytes);
const transitionFadeXml = await transitionFirstZip.file("ppt/slides/slide1.xml").async("string");
const transitionPushXml = await transitionFirstZip.file("ppt/slides/slide2.xml").async("string");
const transitionWipeXml = await transitionFirstZip.file("ppt/slides/slide3.xml").async("string");
assert.match(transitionFadeXml, /<p:transition\b[^>]*\bp14:dur="750"[^>]*>/);
assert.match(transitionFadeXml, /<p:transition\b[^>]*\bspd="med"[^>]*\badvClick="1"[^>]*\badvTm="1250"[^>]*><p:fade \/><\/p:transition>/);
assert.match(transitionPushXml, /<p:transition spd="fast" advClick="0" advTm="0"><p:push dir="r" \/><\/p:transition>/);
assert.match(transitionWipeXml, /<p:transition spd="slow" advClick="1"><p:wipe dir="l" \/><\/p:transition>/);
const transitionImported = await PresentationFile.importPptx(transitionFirstExport);
assert.deepEqual(transitionImported.slides.items[0].transition.toJSON(), { effect: "fade", speed: "medium", durationMs: 750, advanceOnClick: true, advanceAfterMs: 1250 });
assert.deepEqual(transitionImported.slides.items[1].transition.toJSON(), { effect: "push", direction: "right", speed: "fast", advanceOnClick: false, advanceAfterMs: 0 });
assert.deepEqual(transitionImported.slides.items[2].transition.toJSON(), { effect: "wipe", direction: "left", speed: "slow", advanceOnClick: true });
assert.deepEqual(transitionImported.slides.items[0].transition.capability, { sourceBound: true, partPresent: true, editable: true, addable: false });

const transitionVocabularyCases = [
  [{ effect: "blinds", orientation: "vertical" }, /<p:blinds\s+dir="vert"\s*\/>/],
  [{ effect: "checker" }, /<p:checker\s+dir="horz"\s*\/>/],
  [{ effect: "circle" }, /<p:circle\s*\/>/],
  [{ effect: "comb", orientation: "vertical" }, /<p:comb\s+dir="vert"\s*\/>/],
  [{ effect: "cover", direction: "rightUp" }, /<p:cover\s+dir="ru"\s*\/>/],
  [{ effect: "cut", throughBlack: true }, /<p:cut\s+thruBlk="1"\s*\/>/],
  [{ effect: "diamond" }, /<p:diamond\s*\/>/],
  [{ effect: "dissolve" }, /<p:dissolve\s*\/>/],
  [{ effect: "fade", throughBlack: false }, /<p:fade\s+thruBlk="0"\s*\/>/],
  [{ effect: "newsflash" }, /<p:newsflash\s*\/>/],
  [{ effect: "plus" }, /<p:plus\s*\/>/],
  [{ effect: "pull", direction: "leftDown" }, /<p:pull\s+dir="ld"\s*\/>/],
  [{ effect: "push", direction: "up" }, /<p:push\s+dir="u"\s*\/>/],
  [{ effect: "random" }, /<p:random\s*\/>/],
  [{ effect: "randomBar", orientation: "vertical" }, /<p:randomBar\s+dir="vert"\s*\/>/],
  [{ effect: "split", orientation: "horizontal", direction: "in" }, /<p:split\s+orient="horz"\s+dir="in"\s*\/>/],
  [{ effect: "strips", direction: "rightDown" }, /<p:strips\s+dir="rd"\s*\/>/],
  [{ effect: "wedge" }, /<p:wedge\s*\/>/],
  [{ effect: "wheel", spokes: 8 }, /<p:wheel\s+spokes="8"\s*\/>/],
  [{ effect: "wipe", direction: "down" }, /<p:wipe\s+dir="d"\s*\/>/],
  [{ effect: "zoom", direction: "out" }, /<p:zoom\s+dir="out"\s*\/>/],
];
const transitionVocabularyDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
for (const [config] of transitionVocabularyCases) {
  const slide = transitionVocabularyDeck.slides.add({ name: `Transition ${config.effect}` });
  slide.shapes.add({ name: `${config.effect}-title`, text: config.effect });
  slide.setTransition(config);
}
const transitionVocabularyExpected = transitionVocabularyDeck.slides.items.map((slide) => slide.transition.toJSON());
const transitionVocabularyExport = await PresentationFile.exportPptx(transitionVocabularyDeck);
const transitionVocabularyZip = await JSZip.loadAsync(transitionVocabularyExport.bytes);
for (const [index, [, nativeEffect]] of transitionVocabularyCases.entries()) {
  assert.match(await transitionVocabularyZip.file(`ppt/slides/slide${index + 1}.xml`).async("string"), nativeEffect);
}
const transitionVocabularyImported = await PresentationFile.importPptx(transitionVocabularyExport);
assert.deepEqual(
  transitionVocabularyImported.slides.items.map((slide) => slide.transition.toJSON()),
  transitionVocabularyExpected,
);
assert.ok(transitionVocabularyImported.slides.items.every((slide) => slide.transition.capability.editable));
const transitionVocabularySecondImport = await PresentationFile.importPptx(
  await PresentationFile.exportPptx(transitionVocabularyImported),
);
assert.deepEqual(
  transitionVocabularySecondImport.slides.items.map((slide) => slide.transition.toJSON()),
  transitionVocabularyExpected,
);
assert.throws(() => transitionFade.setTransition({ effect: "circle", orientation: "horizontal" }), /does not accept orientation/);
assert.throws(() => transitionFade.setTransition({ effect: "dissolve", throughBlack: true }), /does not accept throughBlack/);
assert.throws(() => transitionFade.setTransition({ effect: "wheel", spokes: 0 }), /1 through 8/);
assert.throws(() => transitionFade.setTransition({ effect: "wheel", spokes: 9 }), /1 through 8/);
assert.throws(() => transitionFade.setTransition({ effect: "wheel", direction: "in" }), /does not accept direction/);
assert.throws(() => transitionFade.setTransition({ effect: "cover", direction: "diagonal" }), /leftUp.*rightDown/);
assert.throws(() => transitionFade.setTransition({ effect: "split", orientation: "diagonal" }), /horizontal or vertical/);
assert.throws(() => transitionFade.setTransition({ effect: "fade", durationMs: -1 }), /durationMs.*0 through 86400000/);
assert.throws(() => transitionFade.setTransition({ effect: "fade", durationMs: 1.5 }), /durationMs.*integer/);
assert.throws(() => transitionFade.setTransition({ effect: "fade", durationMs: 86_400_001 }), /durationMs.*86400000/);
transitionImported.slides.items[0].setTransition({ effect: "wipe", direction: "down", speed: "slow", durationMs: 500, advanceOnClick: true });
const transitionEditedExport = await PresentationFile.exportPptx(transitionImported);
const transitionEdited = await PresentationFile.importPptx(transitionEditedExport);
assert.deepEqual(transitionEdited.slides.items[0].transition.toJSON(), { effect: "wipe", direction: "down", speed: "slow", durationMs: 500, advanceOnClick: true });
transitionEdited.slides.items[1].clearTransition();
const transitionClearedExport = await PresentationFile.exportPptx(transitionEdited);
const transitionCleared = await PresentationFile.importPptx(transitionClearedExport);
assert.equal(transitionCleared.slides.items[1].transition.configured, false);

// The clone is an exact new SlidePart on first export, so a modeled direct
// transition travels with the clone but cannot be changed before reimport.
const transitionCloneSource = await PresentationFile.importPptx(transitionFirstExport);
const transitionClone = transitionCloneSource.slides.items[2].duplicate();
assert.deepEqual(transitionClone.transition.toJSON(), transitionCloneSource.slides.items[2].transition.toJSON());
const transitionCloneExport = await PresentationFile.exportPptx(transitionCloneSource);
const transitionCloneRoundTrip = await PresentationFile.importPptx(transitionCloneExport);
assert.deepEqual(transitionCloneRoundTrip.slides.items[2].transition.toJSON(), transitionCloneRoundTrip.slides.items[3].transition.toJSON());

const transitionAbsentDeck = Presentation.create();
transitionAbsentDeck.slides.add({ name: "No transition" }).shapes.add({ text: "No transition" });
const transitionAbsentPptx = await PresentationFile.exportPptx(transitionAbsentDeck);
const transitionAbsentZip = await JSZip.loadAsync(transitionAbsentPptx.bytes);
const transitionAbsentImported = await PresentationFile.importPptx(transitionAbsentPptx);
assert.deepEqual(transitionAbsentImported.slides.items[0].transition.capability, { sourceBound: true, partPresent: false, editable: false, addable: true });
transitionAbsentImported.slides.items[0].setTransition({ effect: "wipe", direction: "up", speed: "medium", durationMs: 900, advanceOnClick: true });
const transitionAddedPptx = await PresentationFile.exportPptx(transitionAbsentImported);
const transitionAddedZip = await JSZip.loadAsync(transitionAddedPptx.bytes);
assert.deepEqual(Object.keys(transitionAddedZip.files).sort(), Object.keys(transitionAbsentZip.files).sort());
for (const [path, entry] of Object.entries(transitionAbsentZip.files)) {
  if (entry.dir || path === "ppt/slides/slide1.xml") continue;
  assert.deepEqual(
    await transitionAddedZip.file(path).async("uint8array"),
    await transitionAbsentZip.file(path).async("uint8array"),
    `adding an imported transition must preserve ${path} byte-for-byte`,
  );
}
assert.match(await transitionAddedZip.file("ppt/slides/slide1.xml").async("string"), /<p:transition\b[^>]*\bspd="med"[^>]*\bp14:dur="900"[^>]*\badvClick="1"[^>]*><p:wipe\s+dir="u"\s*\/>/);
const transitionAddedImported = await PresentationFile.importPptx(transitionAddedPptx);
assert.deepEqual(transitionAddedImported.slides.items[0].transition.toJSON(), { effect: "wipe", direction: "up", speed: "medium", durationMs: 900, advanceOnClick: true });
assert.deepEqual(transitionAddedImported.slides.items[0].transition.capability, { sourceBound: true, partPresent: true, editable: true, addable: false });

const timedTransitionZip = await JSZip.loadAsync(transitionAbsentPptx.bytes);
timedTransitionZip.file("ppt/slides/slide1.xml", (await timedTransitionZip.file("ppt/slides/slide1.xml").async("string")).replace("</p:sld>", "<p:timing/></p:sld>"));
const timedTransitionImported = await PresentationFile.importPptx(new FileBlob(await timedTransitionZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }));
assert.deepEqual(timedTransitionImported.slides.items[0].transition.capability, { sourceBound: true, partPresent: false, editable: false, addable: false });
assert.throws(
  () => timedTransitionImported.slides.items[0].setTransition({ effect: "fade" }),
  /source-bound/,
);

const opaqueTransitionZip = await JSZip.loadAsync(transitionFirstExport.bytes);
opaqueTransitionZip.file(
  "ppt/slides/slide1.xml",
  transitionFadeXml.replace(/<p:fade\s*\/>/, '<p:cut xmlns:fixture="urn:office-kit:test" fixture:opaque="kept"/>'),
);
const opaqueTransitionBytes = await opaqueTransitionZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const opaqueTransitionFile = new FileBlob(opaqueTransitionBytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
const opaqueTransitionImported = await PresentationFile.importPptx(opaqueTransitionFile);
assert.equal(opaqueTransitionImported.slides.items[0].transition.configured, false);
assert.deepEqual(opaqueTransitionImported.slides.items[0].transition.capability, { sourceBound: true, partPresent: true, editable: false, addable: false });
assert.throws(
  () => opaqueTransitionImported.slides.items[0].setTransition({ effect: "fade" }),
  /source-bound/,
);
const opaqueTransitionPreserved = await PresentationFile.exportPptx(opaqueTransitionImported);
const opaqueTransitionPreservedZip = await JSZip.loadAsync(opaqueTransitionPreserved.bytes);
assert.match(await opaqueTransitionPreservedZip.file("ppt/slides/slide1.xml").async("string"), /<p:cut\s+xmlns:fixture="urn:office-kit:test"\s+fixture:opaque="kept"\s*\/>/);

const nonCanonicalDurationZip = await JSZip.loadAsync(transitionFirstExport.bytes);
nonCanonicalDurationZip.file("ppt/slides/slide1.xml", transitionFadeXml.replace('p14:dur="750"', 'p14:dur="0.75s"'));
const nonCanonicalDurationBytes = await nonCanonicalDurationZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const nonCanonicalDurationImported = await PresentationFile.importPptx(new FileBlob(nonCanonicalDurationBytes, {
  type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
}));
assert.equal(nonCanonicalDurationImported.slides.items[0].transition.configured, false);
assert.deepEqual(nonCanonicalDurationImported.slides.items[0].transition.capability, { sourceBound: true, partPresent: true, editable: false, addable: false });
assert.throws(() => nonCanonicalDurationImported.slides.items[0].setTransition({ effect: "fade" }), /source-bound/);
assert.deepEqual((await PresentationFile.exportPptx(nonCanonicalDurationImported)).bytes, nonCanonicalDurationBytes);
// Negative DrawingML offsets are retained only for an imported opaque,
// source-bound element. New authoring still rejects them instead of widening
// the public source-free layout profile.
const negativeSourceFreeFrame = Presentation.create();
negativeSourceFreeFrame.slides.add().shapes.add({
  name: "negative-source-free-frame",
  position: { left: -1, top: 0, width: 100, height: 40 },
  text: "must fail before export",
});
await assert.rejects(
  () => PresentationFile.exportPptx(negativeSourceFreeFrame),
  (error) => error?.code === "invalid_presentation_frame",
);

// A source-free layout is intentionally a small reusable authoring profile:
// one canonical master, direct-frame title/body text placeholders, and an
// explicit slide binding. It is native PresentationML, not preview-only model
// metadata or a reconstructed imported template graph.
const authoredLayoutPresentation = Presentation.create({
  slideSize: { width: 1280, height: 720 },
  master: {
    name: "Authoring master",
    placeholders: [{
      type: "title",
      index: 0,
      name: "Title",
      text: [{ runs: [{ text: "Master title prompt", link: { customShow: "Layout route", returnToSlide: true } }] }],
      position: { left: 60, top: 44, width: 1160, height: 82 },
      style: { fontSize: 30, bold: true, color: "#0F172A" },
    }],
  },
});
const authoredLayout = authoredLayoutPresentation.layouts.add({ name: "Title and body", type: "titleAndContent" });
authoredLayout.placeholders.add({
  type: "body",
  index: 1,
  name: "Body",
  text: [{ runs: [{ text: "Master body prompt", link: { customShow: "Layout route" } }] }],
  position: { left: 72, top: 154, width: 1136, height: 490 },
  style: { fontSize: 18, color: "#334155" },
});
assert.equal(authoredLayoutPresentation.layouts.getById(authoredLayout.id), authoredLayout);
assert.equal(authoredLayout.placeholders.count, 1);
const authoredLayoutPlaceholderSummary = authoredLayout.placeholders.summary();
assert.deepEqual(authoredLayoutPlaceholderSummary, {
  ownerId: authoredLayout.id,
  count: 1,
  requiredCount: 0,
  types: ["body"],
  items: [{
    id: authoredLayout.placeholders.getItem("body").id,
    name: "Body",
    type: "body",
    idx: 1,
    index: 1,
    required: false,
    hasDirectPosition: true,
    position: { left: 72, top: 154, width: 1136, height: 490 },
  }],
});
authoredLayoutPlaceholderSummary.items[0].position.left = -1;
assert.equal(authoredLayout.placeholders.getItem("body").position.left, 72);
const authoredLayoutSlide = authoredLayoutPresentation.slides.add({ name: "Reusable layout", layout: "Title and body" });
authoredLayoutPresentation.customShows.add("Layout route", [authoredLayoutSlide]);
assert.equal(authoredLayoutSlide.layoutId, authoredLayout.id);
assert.equal(authoredLayoutSlide.placeholders.count, 2);
const materializedPlaceholderCount = authoredLayoutSlide.shapes.items.length;
assert.equal(authoredLayoutSlide.setLayout(authoredLayout), authoredLayoutSlide);
assert.equal(authoredLayoutSlide.layoutId, authoredLayout.id);
assert.equal(authoredLayoutSlide.placeholders.count, 2);
assert.equal(authoredLayoutSlide.shapes.items.length, materializedPlaceholderCount);
const authoredTitle = authoredLayoutSlide.placeholders.getItem("title");
const authoredBody = authoredLayoutSlide.placeholders.getItem(1);
assert.ok(authoredTitle);
assert.ok(authoredBody);
authoredTitle.text.set("OfficeKit layout title");
authoredBody.text.set("A direct-frame body placeholder survives native export and import.");
const authoredLayoutExport = await PresentationFile.exportPptx(authoredLayoutPresentation);
const authoredLayoutZip = await JSZip.loadAsync(new Uint8Array(await authoredLayoutExport.arrayBuffer()));
const authoredMasterXml = await authoredLayoutZip.file("ppt/slideMasters/slideMaster1.xml").async("text");
const authoredLayoutXml = await authoredLayoutZip.file("ppt/slideLayouts/slideLayout1.xml").async("text");
const authoredSlideXml = await authoredLayoutZip.file("ppt/slides/slide1.xml").async("text");
assert.match(authoredMasterXml, /<p:ph[^>]*type="title"[^>]*idx="0"/);
assert.match(authoredMasterXml, /action="ppaction:\/\/customshow\?id=0&amp;return=true"/);
assert.match(authoredLayoutXml, /<p:sldLayout[^>]*type="obj"/);
assert.match(authoredLayoutXml, /<p:ph[^>]*type="body"[^>]*idx="1"/);
assert.match(authoredLayoutXml, /action="ppaction:\/\/customshow\?id=0"/);
assert.match(authoredSlideXml, /<p:ph[^>]*type="title"[^>]*idx="0"/);
assert.match(authoredSlideXml, /<p:ph[^>]*type="body"[^>]*idx="1"/);
const authoredLayoutImported = await PresentationFile.importPptx(authoredLayoutExport);
assert.equal(authoredLayoutImported.master.placeholders.length, 1);
assert.equal(authoredLayoutImported.layouts.items.length, 1);
assert.equal(authoredLayoutImported.layouts.items[0].type, "obj");
assert.equal(authoredLayoutImported.master.placeholders[0].text[0].runs[0].link.customShow, "Layout route");
assert.equal(authoredLayoutImported.layouts.items[0].placeholders[0].text[0].runs[0].link.customShow, "Layout route");
const importedLayoutSlide = authoredLayoutImported.slides.getItem(0);
assert.equal(importedLayoutSlide.layoutId, authoredLayoutImported.layouts.items[0].id);
assert.equal(importedLayoutSlide.placeholders.getItem("title").text.value, "OfficeKit layout title");
assert.equal(importedLayoutSlide.placeholders.getItem("body").text.value, "A direct-frame body placeholder survives native export and import.");
const authoredLayoutRoundTrip = await PresentationFile.exportPptx(authoredLayoutImported);
assert.equal((await PresentationFile.inspectPptx(authoredLayoutRoundTrip)).ok, true);
const importedLayoutTitle = importedLayoutSlide.placeholders.getItem("title");
importedLayoutTitle.text.replace("OfficeKit layout title", "OfficeKit reviewed layout title");
const authoredLayoutTextEdit = await PresentationFile.exportPptx(authoredLayoutImported);
assert.equal(authoredLayoutTextEdit.metadata.editPlan?.schema, "office-kit/pptx-edit-plan/v1");
const authoredLayoutTextEditImported = await PresentationFile.importPptx(authoredLayoutTextEdit);
assert.equal(authoredLayoutTextEditImported.slides.getItem(0).placeholders.getItem("title").text.value, "OfficeKit reviewed layout title");

const guardedLayoutPresentation = Presentation.create();
const firstGuardedLayout = guardedLayoutPresentation.layouts.add({
  name: "First guarded layout",
  type: "title",
  placeholders: [{ type: "title", index: 0, position: { left: 80, top: 72, width: 960, height: 88 } }],
});
const secondGuardedLayout = guardedLayoutPresentation.layouts.add({ name: "Second guarded layout", type: "blank" });
const guardedLayoutSlide = guardedLayoutPresentation.slides.add({ layout: firstGuardedLayout });
assert.throws(
  () => guardedLayoutSlide.setLayout(secondGuardedLayout),
  /already has materialized placeholders.*changing layouts/i,
);
const slideCountBeforeUnknownLayout = guardedLayoutPresentation.slides.count;
assert.throws(
  () => guardedLayoutPresentation.slides.add({ layout: "Missing layout" }),
  /Unknown presentation layout: Missing layout/,
);
assert.equal(guardedLayoutPresentation.slides.count, slideCountBeforeUnknownLayout);

const insertionPresentation = Presentation.create();
const insertionLayout = insertionPresentation.layouts.add({
  name: "Inserted title",
  type: "title",
  placeholders: [{ type: "title", index: 0, position: { left: 88, top: 68, width: 920, height: 92 } }],
});
const insertionFirst = insertionPresentation.slides.add({ name: "First" });
const insertionThird = insertionPresentation.slides.add({ name: "Third" });
const insertionFront = insertionPresentation.slides.insert({ after: null, name: "Front" });
const insertionSecond = insertionPresentation.slides.insert({ after: 0, name: "Second" });
const insertionAfterFirst = insertionPresentation.slides.insert({ after: insertionFirst, name: "After first", layout: insertionLayout });
assert.deepEqual(insertionPresentation.slides.items.map((slide) => slide.name), ["Front", "Second", "First", "After first", "Third"]);
assert.equal(insertionAfterFirst.placeholders.count, 1);
assert.equal(insertionAfterFirst.placeholders.getItem("title").placeholder.layoutId, insertionLayout.id);
const insertionCountBeforeRejectedTarget = insertionPresentation.slides.count;
const foreignSlide = Presentation.create().slides.add({ name: "Foreign" });
assert.throws(
  () => insertionPresentation.slides.insert({ after: foreignSlide, name: "Rejected" }),
  /insertion target must belong to this presentation/i,
);
assert.throws(
  () => insertionPresentation.slides.insert({ after: 99, name: "Rejected" }),
  /after must be an existing Slide, a 0-based slide index, or null/i,
);
assert.equal(insertionPresentation.slides.count, insertionCountBeforeRejectedTarget);
const insertionRoundTrip = await PresentationFile.importPptx(await PresentationFile.exportPptx(insertionPresentation));
assert.deepEqual(insertionRoundTrip.slides.items.map((slide) => slide.name), ["Front", "Second", "First", "After first", "Third"]);

const invalidSourceFreeLayout = Presentation.create();
const invalidSourceFreeSlide = invalidSourceFreeLayout.slides.add();
const invalidLayout = invalidSourceFreeLayout.layouts.add({
  name: "Missing direct placeholder frame",
  type: "title",
  placeholders: [{ type: "title", index: 0 }],
});
invalidSourceFreeSlide.setLayout(invalidLayout);
await assert.rejects(
  () => PresentationFile.exportPptx(invalidSourceFreeLayout),
  /requires a direct position/i,
);

const customGeometryPresentation = Presentation.create({ slideSize: { width: 400, height: 240 } });
const customGeometrySlide = customGeometryPresentation.slides.add({ name: "Custom geometry" });
const customGeometryShape = customGeometrySlide.shapes.add({
  name: "literal-custom-path",
  geometry: "custom",
  position: { left: 20, top: 20, width: 180, height: 120 },
  fill: "#DBEAFE",
  line: { fill: "#2563EB", width: 2 },
  text: "Inset label",
  textRectangle: { left: 24, top: 18, right: 156, bottom: 96 },
  customPaths: [{
    width: 21_600,
    height: 21_600,
    fillMode: "normal",
    stroke: true,
    extrusionAllowed: false,
    commands: [
      { moveTo: { x: 1_000, y: 2_000 } },
      { lineTo: { x: 20_000, y: 2_000 } },
      { quadraticBezTo: { x1: 21_000, y1: 6_000, x: 18_000, y: 12_000 } },
      { cubicBezTo: { x1: 21_000, y1: 6_000, x2: 18_000, y2: 19_000, x: 10_800, y: 20_000 } },
      { arcTo: { widthRadius: 3_000, heightRadius: 4_000, startAngle: 5_400_000, sweepAngle: 21_600_000 } },
      { close: {} },
    ],
  }, {
    width: 21_600,
    height: 21_600,
    fillMode: "none",
    stroke: false,
    extrusionAllowed: true,
    commands: [
      { moveTo: { x: 4_000, y: 4_000 } },
      { lineTo: { x: 17_600, y: 4_000 } },
      { lineTo: { x: 10_800, y: 17_600 } },
      { close: {} },
    ],
  }, {
    width: 21_600,
    height: 21_600,
    commands: [
      { moveTo: { x: 2_000, y: 19_600 } },
      { lineTo: { x: 19_600, y: 19_600 } },
    ],
  }],
});
assert.equal(customGeometryShape.customPaths[0].commands.length, 6);
assert.deepEqual(customGeometryShape.textRectangle, { left: 24, top: 18, right: 156, bottom: 96 });
assert.deepEqual(customGeometryShape.inspectRecord().textRectangle, customGeometryShape.textRectangle);
assert.deepEqual(customGeometryShape.layoutJson().textRectangle, customGeometryShape.textRectangle);
const customGeometrySvg = await (await customGeometrySlide.export()).text();
assert.match(customGeometrySvg, /A 3000 4000 0 0 1 10800 12000 A 3000 4000 0 0 1 10800 20000 Z/);
assert.match(customGeometrySvg, /fill="none" stroke="none"/);
assert.match(customGeometrySvg, /<text x="56" y="70"[^>]*>Inset label<\/text>/);
customGeometryShape.textRectangle.bottom = 30;
const customTextOverflow = customGeometrySlide.validateLayout().issues.find((issue) => issue.id === customGeometryShape.id && issue.type === "textOverflow");
assert.deepEqual(customTextOverflow?.bbox, [44, 38, 132, 12]);
customGeometryShape.textRectangle.bottom = 96;
const customGeometryPptx = await PresentationFile.exportPptx(customGeometryPresentation);
const customGeometryZip = await JSZip.loadAsync(customGeometryPptx.bytes);
const customGeometryXml = await customGeometryZip.file("ppt/slides/slide1.xml").async("text");
assert.match(customGeometryXml, /<a:quadBezTo><a:pt x="21000" y="6000"\s*\/><a:pt x="18000" y="12000"\s*\/><\/a:quadBezTo>/);
assert.match(customGeometryXml, /<a:arcTo wR="3000" hR="4000" stAng="5400000" swAng="21600000"\s*\/>/);
assert.match(customGeometryXml, /<a:gd name="officeKitTextLeft" fmla="\*\/ 228600 w 1714500"\s*\/>/);
assert.match(customGeometryXml, /<a:gd name="officeKitTextTop" fmla="\*\/ 171450 h 1143000"\s*\/>/);
assert.match(customGeometryXml, /<a:gd name="officeKitTextRight" fmla="\*\/ 1485900 w 1714500"\s*\/>/);
assert.match(customGeometryXml, /<a:gd name="officeKitTextBottom" fmla="\*\/ 914400 h 1143000"\s*\/>/);
assert.match(customGeometryXml, /<a:rect l="officeKitTextLeft" t="officeKitTextTop" r="officeKitTextRight" b="officeKitTextBottom"\s*\/>/);
assert.match(customGeometryXml, /<a:path\b(?=[^>]*\bfill="norm")(?=[^>]*\bstroke="(?:1|true)")(?=[^>]*\bextrusionOk="(?:0|false)")[^>]*>/);
assert.match(customGeometryXml, /<a:path\b(?=[^>]*\bfill="none")(?=[^>]*\bstroke="(?:0|false)")(?=[^>]*\bextrusionOk="(?:1|true)")[^>]*>/);
assert.match(customGeometryXml, /<a:path w="21600" h="21600"><a:moveTo><a:pt x="2000" y="19600"\s*\/><\/a:moveTo>/);
const importedCustomGeometry = await PresentationFile.importPptx(customGeometryPptx);
const importedCustomGeometryShape = importedCustomGeometry.slides.getItem(0).shapes.items[0];
assert.equal(importedCustomGeometryShape.customPaths.length, 3);
assert.deepEqual(importedCustomGeometryShape.textRectangle, { left: 24, top: 18, right: 156, bottom: 96 });
const clonedCustomGeometry = await PresentationFile.importPptx(customGeometryPptx);
clonedCustomGeometry.slides.getItem(0).duplicate();
const clonedCustomGeometryRoundTrip = await PresentationFile.importPptx(await PresentationFile.exportPptx(clonedCustomGeometry));
assert.deepEqual(clonedCustomGeometryRoundTrip.slides.getItem(1).shapes.items[0].textRectangle, { left: 24, top: 18, right: 156, bottom: 96 });
assert.deepEqual(
  {
    fillMode: importedCustomGeometryShape.customPaths[0].fillMode,
    stroke: importedCustomGeometryShape.customPaths[0].stroke,
    extrusionAllowed: importedCustomGeometryShape.customPaths[0].extrusionAllowed,
  },
  { fillMode: "normal", stroke: true, extrusionAllowed: false },
);
assert.deepEqual(
  {
    fillMode: importedCustomGeometryShape.customPaths[1].fillMode,
    stroke: importedCustomGeometryShape.customPaths[1].stroke,
    extrusionAllowed: importedCustomGeometryShape.customPaths[1].extrusionAllowed,
  },
  { fillMode: "none", stroke: false, extrusionAllowed: true },
);
assert.equal(Object.hasOwn(importedCustomGeometryShape.customPaths[2], "fillMode"), false);
assert.equal(Object.hasOwn(importedCustomGeometryShape.customPaths[2], "stroke"), false);
assert.equal(Object.hasOwn(importedCustomGeometryShape.customPaths[2], "extrusionAllowed"), false);
assert.deepEqual(importedCustomGeometryShape.customPaths[0].commands[2], {
  quadraticBezTo: { x1: 21_000, y1: 6_000, x: 18_000, y: 12_000 },
});
assert.deepEqual(importedCustomGeometryShape.customPaths[0].commands[4], {
  arcTo: { widthRadius: 3_000, heightRadius: 4_000, startAngle: 5_400_000, sweepAngle: 21_600_000 },
});
importedCustomGeometryShape.customPaths[0].commands[2].quadraticBezTo.x1 = 20_500;
importedCustomGeometryShape.customPaths[0].commands[4].arcTo.sweepAngle = -10_800_000;
importedCustomGeometryShape.customPaths[0].fillMode = "none";
delete importedCustomGeometryShape.customPaths[0].stroke;
importedCustomGeometryShape.customPaths[0].extrusionAllowed = true;
importedCustomGeometryShape.customPaths[1].fillMode = "normal";
importedCustomGeometryShape.customPaths[1].stroke = true;
delete importedCustomGeometryShape.customPaths[1].extrusionAllowed;
importedCustomGeometryShape.textRectangle.right = 160;
importedCustomGeometryShape.textRectangle.bottom = 100;
const editedCustomGeometry = await PresentationFile.importPptx(await PresentationFile.exportPptx(importedCustomGeometry));
assert.equal(editedCustomGeometry.slides.getItem(0).shapes.items[0].customPaths[0].commands[2].quadraticBezTo.x1, 20_500);
assert.equal(editedCustomGeometry.slides.getItem(0).shapes.items[0].customPaths[0].commands[4].arcTo.sweepAngle, -10_800_000);
assert.deepEqual(editedCustomGeometry.slides.getItem(0).shapes.items[0].textRectangle, { left: 24, top: 18, right: 160, bottom: 100 });
assert.deepEqual(
  editedCustomGeometry.slides.getItem(0).shapes.items[0].customPaths.map((path) => ({
    fillMode: path.fillMode,
    strokePresent: Object.hasOwn(path, "stroke"),
    stroke: path.stroke,
    extrusionAllowedPresent: Object.hasOwn(path, "extrusionAllowed"),
    extrusionAllowed: path.extrusionAllowed,
  })),
  [
    { fillMode: "none", strokePresent: false, stroke: undefined, extrusionAllowedPresent: true, extrusionAllowed: true },
    { fillMode: "normal", strokePresent: true, stroke: true, extrusionAllowedPresent: false, extrusionAllowed: undefined },
    { fillMode: undefined, strokePresent: false, stroke: undefined, extrusionAllowedPresent: false, extrusionAllowed: undefined },
  ],
);
assert.match(await (await editedCustomGeometry.slides.getItem(0).export()).text(), /A 3000 4000 0 0 0 10800 12000 Z/);

const formulaGeometryConfig = (name, position) => ({
  name,
  geometry: "custom",
  position,
  fill: "#DCFCE7",
  line: { fill: "#15803D", width: 2 },
  text: "Formula shape",
  textRectangle: { left: "x1", top: "t", right: "x2", bottom: "b" },
  customAdjustments: [
    { name: "adjX", formula: "val 25000" },
    { name: "adjY", formula: "val 50000" },
    { name: "adjRadius", formula: "val 250000" },
    { name: "adjSweep", formula: "val 10800000" },
  ],
  customGuides: [
    { name: "x1", formula: "*/ w adjX 100000" },
    { name: "y1", formula: "*/ h 1 2" },
    { name: "x2", formula: "+- r 0 x1" },
    { name: "radius", formula: "min wd4 hd4" },
    { name: "zeroAngle", formula: "+- cd4 0 cd4" },
    { name: "addDiv", formula: "+/ w 0 2" },
    { name: "conditional", formula: "?: adjX w h" },
    { name: "absolute", formula: "abs -5" },
    { name: "arcTangent", formula: "at2 1 1" },
    { name: "cosArcTangent", formula: "cat2 100 1 1" },
    { name: "cosine", formula: "cos 100 cd4" },
    { name: "maximum", formula: "max w h" },
    { name: "modulus", formula: "mod 3 4 12" },
    { name: "pinned", formula: "pin 0 adjX 100000" },
    { name: "sinArcTangent", formula: "sat2 100 1 1" },
    { name: "sine", formula: "sin 100 cd4" },
    { name: "squareRoot", formula: "sqrt 144" },
    { name: "tangent", formula: "tan 100 cd8" },
  ],
  customConnectionSites: [
    { angle: 180, x: "x1", y: "y1" },
    { angle: "zeroAngle", x: "x2", y: "y1" },
  ],
  customAdjustmentHandles: [
    { kind: "xy", xAdjustment: "adjX", minX: 0, maxX: 100000, yAdjustment: "adjY", minY: 0, maxY: 100000, x: "x1", y: "y1" },
    { kind: "polar", radialAdjustment: "adjRadius", minRadius: 0, maxRadius: 500000, angleAdjustment: "adjSweep", minAngle: 0, maxAngle: 360, x: "x2", y: "y1" },
  ],
  customPaths: [{
    width: Math.round(position.width * 9_525),
    height: Math.round(position.height * 9_525),
    commands: [
      { moveTo: { x: "x1", y: "y1" } },
      { lineTo: { x: "x2", y: "y1" } },
      { arcTo: { widthRadius: "radius", heightRadius: "radius", startAngle: "zeroAngle", sweepAngle: "adjSweep" } },
      { close: {} },
    ],
  }],
});
const formulaGeometryPresentation = Presentation.create({ slideSize: { width: 500, height: 300 } });
const formulaGeometrySlide = formulaGeometryPresentation.slides.add({ name: "Formula custom geometry" });
const formulaGeometryShape = formulaGeometrySlide.shapes.add(formulaGeometryConfig(
  "formula-custom-path",
  { left: 20, top: 20, width: 200, height: 100 },
));
const formulaGeometryGroup = formulaGeometrySlide.groups.add({
  name: "Formula group",
  position: { left: 260, top: 20, width: 200, height: 140 },
  childFrame: { left: 0, top: 0, width: 200, height: 140 },
});
formulaGeometryGroup.shapes.add(formulaGeometryConfig(
  "grouped-formula-custom-path",
  { left: 20, top: 20, width: 160, height: 100 },
));
const builtinGeometryShape = formulaGeometrySlide.shapes.add({
  name: "builtin-guide-default-path",
  geometry: "custom",
  position: { left: 20, top: 170, width: 200, height: 100 },
  fill: "#FDE68A",
  line: { fill: "#B45309", width: 2 },
  text: "Built-in guides",
  textRectangle: { left: "l", top: "t", right: "r", bottom: "b" },
  customAdjustments: [
    { name: "adjX", formula: "val 25000" },
    { name: "adjY", formula: "val 50000" },
    { name: "adjRadius", formula: "val 250000" },
    { name: "adjSweep", formula: "val 10800000" },
  ],
  customConnectionSites: [
    { angle: "3cd4", x: "hc", y: "t" },
    { angle: "cd2", x: "l", y: "vc" },
    { angle: "cd4", x: "hc", y: "b" },
    { angle: 0, x: "r", y: "vc" },
  ],
  customAdjustmentHandles: [
    { kind: "xy", xAdjustment: "adjX", minX: "l", maxX: "r", yAdjustment: "adjY", minY: "t", maxY: "b", x: "hc", y: "vc" },
    { kind: "polar", radialAdjustment: "adjRadius", minRadius: "l", maxRadius: "ss", angleAdjustment: "adjSweep", minAngle: "t", maxAngle: "cd2", x: "r", y: "vc" },
  ],
  customPaths: [{
    commands: [
      { moveTo: { x: "l", y: "vc" } },
      { lineTo: { x: "hc", y: "t" } },
      { arcTo: { widthRadius: "wd4", heightRadius: "hd4", startAngle: "3cd4", sweepAngle: "cd2" } },
      { lineTo: { x: "r", y: "vc" } },
      { lineTo: { x: "hc", y: "b" } },
      { close: {} },
    ],
  }, {
    width: 100,
    fillMode: "none",
    stroke: false,
    commands: [
      { moveTo: { x: 0, y: 0 } },
      { lineTo: { x: 100, y: 0 } },
    ],
  }],
});
assert.deepEqual(formulaGeometryShape.inspectRecord().customAdjustmentCount, 4);
assert.deepEqual(formulaGeometryShape.inspectRecord().customGuideCount, 18);
assert.deepEqual(formulaGeometryShape.inspectRecord().customConnectionSiteCount, 2);
assert.deepEqual(formulaGeometryShape.inspectRecord().customAdjustmentHandleCount, 2);
assert.deepEqual(formulaGeometryShape.layoutJson().customConnectionSites, [
  { angle: 180, x: "x1", y: "y1" },
  { angle: "zeroAngle", x: "x2", y: "y1" },
]);
assert.deepEqual(formulaGeometryShape.layoutJson().customAdjustmentHandles, [
  { kind: "xy", xAdjustment: "adjX", minX: 0, maxX: 100000, yAdjustment: "adjY", minY: 0, maxY: 100000, x: "x1", y: "y1" },
  { kind: "polar", radialAdjustment: "adjRadius", minRadius: 0, maxRadius: 500000, angleAdjustment: "adjSweep", minAngle: 0, maxAngle: 360, x: "x2", y: "y1" },
]);
assert.deepEqual(formulaGeometryShape.textRectangle, { left: "x1", top: "t", right: "x2", bottom: "b" });
assert.deepEqual(formulaGeometryShape.textFrame(), { left: 70, top: 20, width: 100, height: 100 });
assert.throws(
  () => formulaGeometrySlide.shapes.getConnectionSiteIndex(formulaGeometryShape, "right"),
  /requires an explicit connection-site index/,
);
assert.match(formulaGeometryShape.toSvg(), /M 476250 476250 L 1428750 476250 A 238125 238125/);
assert.match(builtinGeometryShape.toSvg(), /M 0 476250 L 952500 0 A 476250 238125/);
assert.equal(Object.hasOwn(builtinGeometryShape.customPaths[0], "width"), false);
assert.equal(Object.hasOwn(builtinGeometryShape.customPaths[0], "height"), false);
assert.equal(builtinGeometryShape.customPaths[1].width, 100);
assert.equal(Object.hasOwn(builtinGeometryShape.customPaths[1], "height"), false);
assert.deepEqual(formulaGeometryShape.layoutJson().customGuides[0], { name: "x1", formula: "*/ w adjX 100000" });
const formulaGeometryPptx = await PresentationFile.exportPptx(formulaGeometryPresentation);
const formulaGeometryZip = await JSZip.loadAsync(formulaGeometryPptx.bytes);
const formulaGeometryXml = await formulaGeometryZip.file("ppt/slides/slide1.xml").async("text");
assert.match(formulaGeometryXml, /<a:avLst><a:gd name="adjX" fmla="val 25000"\s*\/><a:gd name="adjY" fmla="val 50000"\s*\/><a:gd name="adjRadius" fmla="val 250000"\s*\/><a:gd name="adjSweep" fmla="val 10800000"\s*\/><\/a:avLst>/);
assert.match(formulaGeometryXml, /<a:gd name="x1" fmla="\*\/ w adjX 100000"\s*\/>/);
assert.match(formulaGeometryXml, /<a:gd name="tangent" fmla="tan 100 cd8"\s*\/>/);
assert.match(formulaGeometryXml, /<a:ahLst><a:ahXY gdRefX="adjX" minX="0" maxX="100000" gdRefY="adjY" minY="0" maxY="100000"><a:pos x="x1" y="y1"\s*\/><\/a:ahXY><a:ahPolar gdRefR="adjRadius" minR="0" maxR="500000" gdRefAng="adjSweep" minAng="0" maxAng="21600000"><a:pos x="x2" y="y1"\s*\/><\/a:ahPolar><\/a:ahLst>/);
assert.match(formulaGeometryXml, /<a:cxnLst><a:cxn ang="10800000"><a:pos x="x1" y="y1"\s*\/><\/a:cxn><a:cxn ang="zeroAngle"><a:pos x="x2" y="y1"\s*\/><\/a:cxn><\/a:cxnLst>/);
assert.match(formulaGeometryXml, /<a:pt x="x1" y="y1"\s*\/>/);
assert.match(formulaGeometryXml, /<a:arcTo wR="radius" hR="radius" stAng="zeroAngle" swAng="adjSweep"\s*\/>/);
assert.match(formulaGeometryXml, /<a:rect l="x1" t="t" r="x2" b="b"\s*\/>/);
assert.match(formulaGeometryXml, /<a:ahXY gdRefX="adjX" minX="l" maxX="r" gdRefY="adjY" minY="t" maxY="b"><a:pos x="hc" y="vc"\s*\/>/);
assert.match(formulaGeometryXml, /<a:cxn ang="3cd4"><a:pos x="hc" y="t"\s*\/>/);
assert.match(formulaGeometryXml, /<a:path><a:moveTo><a:pt x="l" y="vc"\s*\/>/);
assert.match(formulaGeometryXml, /<a:arcTo wR="wd4" hR="hd4" stAng="3cd4" swAng="cd2"\s*\/>/);
assert.match(formulaGeometryXml, /<a:path w="100" fill="none" stroke="0"><a:moveTo>/);
assert.doesNotMatch(formulaGeometryXml, /officeKitText(?:Left|Top|Right|Bottom)/);
const importedFormulaGeometry = await PresentationFile.importPptx(formulaGeometryPptx);
const importedFormulaShape = importedFormulaGeometry.slides.getItem(0).shapes.items[0];
const importedBuiltinGeometryShape = itemByName(importedFormulaGeometry.slides.getItem(0).shapes.items, "builtin-guide-default-path");
assert.equal(importedFormulaShape.customAdjustments.length, 4);
assert.equal(importedFormulaShape.customGuides.length, 18);
assert.deepEqual(importedFormulaShape.customConnectionSites, [
  { angle: 180, x: "x1", y: "y1" },
  { angle: "zeroAngle", x: "x2", y: "y1" },
]);
assert.deepEqual(importedFormulaShape.customAdjustmentHandles, [
  { kind: "xy", xAdjustment: "adjX", minX: 0, maxX: 100000, yAdjustment: "adjY", minY: 0, maxY: 100000, x: "x1", y: "y1" },
  { kind: "polar", radialAdjustment: "adjRadius", minRadius: 0, maxRadius: 500000, angleAdjustment: "adjSweep", minAngle: 0, maxAngle: 360, x: "x2", y: "y1" },
]);
assert.deepEqual(importedFormulaShape.textRectangle, { left: "x1", top: "t", right: "x2", bottom: "b" });
assert.equal(importedFormulaShape.customPaths[0].commands[0].moveTo.x, "x1");
assert.equal(importedFormulaShape.customPaths[0].commands[2].arcTo.widthRadius, "radius");
assert.equal(Object.hasOwn(importedBuiltinGeometryShape.customPaths[0], "width"), false);
assert.equal(Object.hasOwn(importedBuiltinGeometryShape.customPaths[0], "height"), false);
assert.equal(importedBuiltinGeometryShape.customPaths[1].width, 100);
assert.equal(Object.hasOwn(importedBuiltinGeometryShape.customPaths[1], "height"), false);
assert.deepEqual(importedBuiltinGeometryShape.customConnectionSites[0], { angle: "3cd4", x: "hc", y: "t" });
assert.deepEqual(importedBuiltinGeometryShape.customAdjustmentHandles[0], {
  kind: "xy", xAdjustment: "adjX", minX: "l", maxX: "r", yAdjustment: "adjY", minY: "t", maxY: "b", x: "hc", y: "vc",
});
assert.equal(importedBuiltinGeometryShape.customPaths[0].commands[2].arcTo.widthRadius, "wd4");
importedBuiltinGeometryShape.customPaths[0].commands[1].lineTo.x = "r";
assert.equal(importedFormulaGeometry.slides.getItem(0).groups.items[0].shapes.items[0].customGuides.length, 18);
assert.equal(importedFormulaGeometry.slides.getItem(0).groups.items[0].shapes.items[0].customConnectionSites.length, 2);
assert.equal(importedFormulaGeometry.slides.getItem(0).groups.items[0].shapes.items[0].customAdjustmentHandles.length, 2);
const unknownTextRectangleGuideXml = formulaGeometryXml.replace('r="x2"', 'r="missingRectGuide"');
assert.notEqual(unknownTextRectangleGuideXml, formulaGeometryXml);
const unknownTextRectangleGuideFile = await PresentationFile.patchPptx(formulaGeometryPptx, [{ path: "ppt/slides/slide1.xml", xml: unknownTextRectangleGuideXml }]);
const unknownTextRectangleGuidePresentation = await PresentationFile.importPptx(unknownTextRectangleGuideFile);
const opaqueUnknownTextRectangleGuide = itemByName(unknownTextRectangleGuidePresentation.slides.getItem(0).shapes.items, "formula-custom-path");
assert.equal(opaqueUnknownTextRectangleGuide.customPaths.length, 0);
assert.equal(opaqueUnknownTextRectangleGuide.textRectangle, undefined);
const preservedUnknownTextRectangleGuide = await PresentationFile.exportPptx(unknownTextRectangleGuidePresentation);
const preservedUnknownTextRectangleGuideZip = await JSZip.loadAsync(preservedUnknownTextRectangleGuide.bytes);
assert.match(await preservedUnknownTextRectangleGuideZip.file("ppt/slides/slide1.xml").async("text"), /<a:rect l="x1" t="t" r="missingRectGuide" b="b"\s*\/>/);
const invertedTextRectangleGuideXml = formulaGeometryXml.replace('<a:rect l="x1" t="t" r="x2" b="b" />', '<a:rect l="x2" t="t" r="x1" b="b" />');
assert.notEqual(invertedTextRectangleGuideXml, formulaGeometryXml);
const invertedTextRectangleGuideFile = await PresentationFile.patchPptx(formulaGeometryPptx, [{ path: "ppt/slides/slide1.xml", xml: invertedTextRectangleGuideXml }]);
const invertedTextRectangleGuidePresentation = await PresentationFile.importPptx(invertedTextRectangleGuideFile);
assert.equal(itemByName(invertedTextRectangleGuidePresentation.slides.getItem(0).shapes.items, "formula-custom-path").customPaths.length, 0);
const emptyFormulaTopologyXml = customGeometryXml.replace("</a:gdLst><a:rect", "</a:gdLst><a:ahLst /><a:cxnLst /><a:rect");
assert.notEqual(emptyFormulaTopologyXml, customGeometryXml);
const emptyFormulaTopologyFile = await PresentationFile.patchPptx(customGeometryPptx, [{ path: "ppt/slides/slide1.xml", xml: emptyFormulaTopologyXml }]);
const importedEmptyFormulaTopology = await PresentationFile.importPptx(emptyFormulaTopologyFile);
assert.equal(importedEmptyFormulaTopology.slides.getItem(0).shapes.items[0].customPaths.length, 3);
const invalidHandleFormulaTopologyXml = formulaGeometryXml.replace('<a:pos x="x1" y="y1" />', '<a:pos x="x1" y="y1" data="unexpected" />');
assert.notEqual(invalidHandleFormulaTopologyXml, formulaGeometryXml);
const invalidHandleFormulaTopologyFile = await PresentationFile.patchPptx(formulaGeometryPptx, [{ path: "ppt/slides/slide1.xml", xml: invalidHandleFormulaTopologyXml }]);
const importedInvalidHandleFormulaTopology = await PresentationFile.importPptx(invalidHandleFormulaTopologyFile);
const opaqueHandleFormulaShape = itemByName(importedInvalidHandleFormulaTopology.slides.getItem(0).shapes.items, "formula-custom-path");
assert.equal(opaqueHandleFormulaShape.customPaths.length, 0);
assert.equal(opaqueHandleFormulaShape.customAdjustmentHandles.length, 0);
const preservedInvalidHandleFormulaTopology = await PresentationFile.exportPptx(importedInvalidHandleFormulaTopology);
const preservedInvalidHandleFormulaZip = await JSZip.loadAsync(preservedInvalidHandleFormulaTopology.bytes);
assert.match(await preservedInvalidHandleFormulaZip.file("ppt/slides/slide1.xml").async("text"), /<a:pos x="x1" y="y1" data="unexpected"\s*\/>/);
const invalidConnectionSiteXml = formulaGeometryXml.replace(
  '<a:pos x="x1" y="y1" />',
  '<a:pos x="x1" y="y1" data="unexpected" />',
);
assert.notEqual(invalidConnectionSiteXml, formulaGeometryXml);
const invalidConnectionSiteFile = await PresentationFile.patchPptx(formulaGeometryPptx, [{ path: "ppt/slides/slide1.xml", xml: invalidConnectionSiteXml }]);
const invalidConnectionSitePresentation = await PresentationFile.importPptx(invalidConnectionSiteFile);
const opaqueConnectionSiteShape = itemByName(invalidConnectionSitePresentation.slides.getItem(0).shapes.items, "formula-custom-path");
assert.equal(opaqueConnectionSiteShape.customPaths.length, 0);
assert.equal(opaqueConnectionSiteShape.customConnectionSites.length, 0);
const preservedInvalidConnectionSite = await PresentationFile.exportPptx(invalidConnectionSitePresentation);
const preservedInvalidConnectionSiteZip = await JSZip.loadAsync(preservedInvalidConnectionSite.bytes);
assert.match(await preservedInvalidConnectionSiteZip.file("ppt/slides/slide1.xml").async("text"), /<a:pos x="x1" y="y1" data="unexpected"\s*\/>/);
importedFormulaShape.customAdjustments[0].formula = "val 30000";
assert.match(importedFormulaShape.toSvg(), /M 571500 476250 L 1333500 476250/);
importedFormulaShape.customConnectionSites[0].angle = 90;
importedFormulaShape.customAdjustmentHandles[0].maxX = 90000;
importedFormulaShape.customAdjustmentHandles[1].x = "x1";
importedFormulaShape.textRectangle.right = 180;
const editedFormulaGeometry = await PresentationFile.importPptx(await PresentationFile.exportPptx(importedFormulaGeometry));
assert.equal(editedFormulaGeometry.slides.getItem(0).shapes.items[0].customAdjustments[0].formula, "val 30000");
assert.equal(editedFormulaGeometry.slides.getItem(0).shapes.items[0].customConnectionSites[0].angle, 90);
assert.equal(editedFormulaGeometry.slides.getItem(0).shapes.items[0].customAdjustmentHandles[0].maxX, 90000);
assert.equal(editedFormulaGeometry.slides.getItem(0).shapes.items[0].customAdjustmentHandles[1].x, "x1");
assert.equal(editedFormulaGeometry.slides.getItem(0).shapes.items[0].customPaths[0].commands[0].moveTo.x, "x1");
const editedBuiltinGeometryShape = itemByName(editedFormulaGeometry.slides.getItem(0).shapes.items, "builtin-guide-default-path");
assert.equal(editedBuiltinGeometryShape.customPaths[0].commands[1].lineTo.x, "r");
assert.equal(Object.hasOwn(editedBuiltinGeometryShape.customPaths[0], "width"), false);
assert.equal(Object.hasOwn(editedBuiltinGeometryShape.customPaths[0], "height"), false);
assert.equal(editedBuiltinGeometryShape.customPaths[1].width, 100);
assert.equal(Object.hasOwn(editedBuiltinGeometryShape.customPaths[1], "height"), false);
assert.deepEqual(editedFormulaGeometry.slides.getItem(0).shapes.items[0].textRectangle, { left: "x1", top: "t", right: 180, bottom: "b" });
const editedFormulaGeometryZip = await JSZip.loadAsync((await PresentationFile.exportPptx(importedFormulaGeometry)).bytes);
const editedFormulaGeometryXml = await editedFormulaGeometryZip.file("ppt/slides/slide1.xml").async("text");
assert.match(editedFormulaGeometryXml, /<a:gd name="officeKitTextRight" fmla="\*\/ 1714500 w 1905000"\s*\/>/);
assert.match(editedFormulaGeometryXml, /<a:rect l="x1" t="t" r="officeKitTextRight" b="b"\s*\/>/);
const changedConnectionSiteTopology = await PresentationFile.importPptx(formulaGeometryPptx);
changedConnectionSiteTopology.slides.getItem(0).shapes.items[0].customConnectionSites.pop();
await assert.rejects(
  () => PresentationFile.exportPptx(changedConnectionSiteTopology),
  (error) => error?.code === "unsupported_presentation_edit" && /connection-site list length/i.test(error.message),
);
const changedAdjustmentHandleTopology = await PresentationFile.importPptx(formulaGeometryPptx);
changedAdjustmentHandleTopology.slides.getItem(0).shapes.items[0].customAdjustmentHandles[0].xAdjustment = "adjY";
await assert.rejects(
  () => PresentationFile.exportPptx(changedAdjustmentHandleTopology),
  (error) => error?.code === "unsupported_presentation_edit" && /adjustment-handle order, kind, and controlled adjustment identity/i.test(error.message),
);

const customSiteConnectorDeck = Presentation.create({ slideSize: { width: 500, height: 300 } });
const customSiteConnectorSlide = customSiteConnectorDeck.slides.add({ name: "Custom connection sites" });
const customSiteConnectorSource = customSiteConnectorSlide.shapes.add(formulaGeometryConfig(
  "custom-site-source",
  { left: 20, top: 20, width: 200, height: 100 },
));
const customSiteConnectorTarget = customSiteConnectorSlide.shapes.add({
  name: "custom-site-target",
  geometry: "rect",
  position: { left: 300, top: 210, width: 140, height: 60 },
  text: "Target",
});
const customSiteConnector = customSiteConnectorSlide.shapes.connect(customSiteConnectorSource, customSiteConnectorTarget, {
  name: "formula-site-connector",
  fromIdx: 1,
  toIdx: 1,
});
assert.deepEqual(customSiteConnector.start, { x: 170, y: 70 });
customSiteConnectorSource.customAdjustments[0].formula = "val 30000";
assert.deepEqual(customSiteConnector.start, { x: 160, y: 70 });
const customSiteConnectorRoundTrip = await PresentationFile.importPptx(await PresentationFile.exportPptx(customSiteConnectorDeck));
const roundTripCustomSiteConnector = itemByName(customSiteConnectorRoundTrip.slides.getItem(0).connectors.items, "formula-site-connector");
assert.equal(roundTripCustomSiteConnector.startSiteIndex, 1);
assert.equal(roundTripCustomSiteConnector.endSiteIndex, 1);
assert.deepEqual(roundTripCustomSiteConnector.start, { x: 160, y: 70 });

const defaultTextRectanglePresentation = Presentation.create({ slideSize: { width: 160, height: 100 } });
const defaultTextRectangleSlide = defaultTextRectanglePresentation.slides.add({ name: "Default custom text bounds" });
defaultTextRectangleSlide.shapes.add({
  geometry: "custom",
  position: { left: 10, top: 10, width: 120, height: 70 },
  text: "Full shape bounds",
  customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 0, y: 0 } }, { lineTo: { x: 100, y: 100 } }] }],
});
const defaultTextRectanglePptx = await PresentationFile.exportPptx(defaultTextRectanglePresentation);
const defaultTextRectangleZip = await JSZip.loadAsync(defaultTextRectanglePptx.bytes);
assert.doesNotMatch(await defaultTextRectangleZip.file("ppt/slides/slide1.xml").async("text"), /<a:rect\b/);
const importedDefaultTextRectangle = await PresentationFile.importPptx(defaultTextRectanglePptx);
const importedDefaultTextRectangleShape = importedDefaultTextRectangle.slides.getItem(0).shapes.items[0];
assert.equal(importedDefaultTextRectangleShape.textRectangle, undefined);
importedDefaultTextRectangleShape.textRectangle = { left: -4, top: 5, right: 124, bottom: 64 };
const addedTextRectangle = await PresentationFile.importPptx(await PresentationFile.exportPptx(importedDefaultTextRectangle));
assert.deepEqual(addedTextRectangle.slides.getItem(0).shapes.items[0].textRectangle, { left: -4, top: 5, right: 124, bottom: 64 });
addedTextRectangle.slides.getItem(0).shapes.items[0].textRectangle = undefined;
const removedTextRectangleFile = await PresentationFile.exportPptx(addedTextRectangle);
const removedTextRectangleZip = await JSZip.loadAsync(removedTextRectangleFile.bytes);
assert.doesNotMatch(await removedTextRectangleZip.file("ppt/slides/slide1.xml").async("text"), /<a:rect\b/);
assert.equal((await PresentationFile.importPptx(removedTextRectangleFile)).slides.getItem(0).shapes.items[0].textRectangle, undefined);
const shadedCustomGeometryXml = customGeometryXml.replace('fill="norm"', 'fill="lighten"');
assert.notEqual(shadedCustomGeometryXml, customGeometryXml);
const shadedCustomGeometryFile = await PresentationFile.patchPptx(customGeometryPptx, [{ path: "ppt/slides/slide1.xml", xml: shadedCustomGeometryXml }]);
const shadedCustomGeometry = await PresentationFile.importPptx(shadedCustomGeometryFile);
const opaqueShadedGeometry = itemByName(shadedCustomGeometry.slides.getItem(0).shapes.items, "literal-custom-path");
assert.equal(opaqueShadedGeometry.customPaths.length, 0);
const preservedShadedGeometry = await PresentationFile.exportPptx(shadedCustomGeometry);
const preservedShadedGeometryZip = await JSZip.loadAsync(preservedShadedGeometry.bytes);
assert.match(await preservedShadedGeometryZip.file("ppt/slides/slide1.xml").async("text"), /<a:path\b[^>]*\bfill="lighten"/);
opaqueShadedGeometry.name = "unsafe-shaded-geometry-edit";
await assert.rejects(
  () => PresentationFile.exportPptx(shadedCustomGeometry),
  (error) => error?.code === "unsupported_presentation_edit",
);
const invalidStrokeCustomGeometryXml = customGeometryXml.replace(/stroke="(?:1|true)"/, 'stroke="maybe"');
assert.notEqual(invalidStrokeCustomGeometryXml, customGeometryXml);
const invalidStrokeCustomGeometryFile = await PresentationFile.patchPptx(customGeometryPptx, [{ path: "ppt/slides/slide1.xml", xml: invalidStrokeCustomGeometryXml }]);
const invalidStrokeCustomGeometry = await PresentationFile.importPptx(invalidStrokeCustomGeometryFile);
assert.equal(itemByName(invalidStrokeCustomGeometry.slides.getItem(0).shapes.items, "literal-custom-path").customPaths.length, 0);
const preservedInvalidStroke = await PresentationFile.exportPptx(invalidStrokeCustomGeometry);
const preservedInvalidStrokeZip = await JSZip.loadAsync(preservedInvalidStroke.bytes);
assert.match(await preservedInvalidStrokeZip.file("ppt/slides/slide1.xml").async("text"), /<a:path\b[^>]*\bstroke="maybe"/);
const builtinFormulaCustomGeometryXml = customGeometryXml.replace('wR="3000"', 'wR="wd2"');
assert.notEqual(builtinFormulaCustomGeometryXml, customGeometryXml);
const builtinFormulaCustomGeometryFile = await PresentationFile.patchPptx(customGeometryPptx, [{ path: "ppt/slides/slide1.xml", xml: builtinFormulaCustomGeometryXml }]);
const builtinFormulaCustomGeometry = await PresentationFile.importPptx(builtinFormulaCustomGeometryFile);
const builtinFormulaCustomGeometryShape = itemByName(builtinFormulaCustomGeometry.slides.getItem(0).shapes.items, "literal-custom-path");
assert.equal(builtinFormulaCustomGeometryShape.customPaths.length, 3);
assert.equal(builtinFormulaCustomGeometryShape.customPaths[0].commands[4].arcTo.widthRadius, "wd2");
const preservedBuiltinFormulaGeometry = await PresentationFile.exportPptx(builtinFormulaCustomGeometry);
const preservedBuiltinFormulaGeometryZip = await JSZip.loadAsync(preservedBuiltinFormulaGeometry.bytes);
assert.match(await preservedBuiltinFormulaGeometryZip.file("ppt/slides/slide1.xml").async("text"), /<a:arcTo wR="wd2" hR="4000" stAng="5400000" swAng="21600000"\s*\/>/);
const unknownFormulaCustomGeometryXml = customGeometryXml.replace('wR="3000"', 'wR="missingGuide"');
const unknownFormulaCustomGeometryFile = await PresentationFile.patchPptx(customGeometryPptx, [{ path: "ppt/slides/slide1.xml", xml: unknownFormulaCustomGeometryXml }]);
const unknownFormulaCustomGeometry = await PresentationFile.importPptx(unknownFormulaCustomGeometryFile);
const opaqueFormulaGeometry = itemByName(unknownFormulaCustomGeometry.slides.getItem(0).shapes.items, "literal-custom-path");
assert.equal(opaqueFormulaGeometry.customPaths.length, 0);
const preservedUnknownFormulaGeometry = await PresentationFile.exportPptx(unknownFormulaCustomGeometry);
const preservedUnknownFormulaGeometryZip = await JSZip.loadAsync(preservedUnknownFormulaGeometry.bytes);
assert.match(await preservedUnknownFormulaGeometryZip.file("ppt/slides/slide1.xml").async("text"), /<a:arcTo wR="missingGuide" hR="4000" stAng="5400000" swAng="21600000"\s*\/>/);
opaqueFormulaGeometry.name = "unsafe-formula-geometry-edit";
await assert.rejects(
  () => PresentationFile.exportPptx(unknownFormulaCustomGeometry),
  (error) => error?.code === "unsupported_presentation_edit",
);
const officeKitTextGuideList = /<a:gdLst><a:gd name="officeKitTextLeft"[\s\S]*?<\/a:gdLst>/;
assert.match(customGeometryXml, officeKitTextGuideList);
const literalTextRectangleXml = customGeometryXml
  .replace(officeKitTextGuideList, "")
  .replace(
    '<a:rect l="officeKitTextLeft" t="officeKitTextTop" r="officeKitTextRight" b="officeKitTextBottom" />',
    '<a:rect l="228600" t="171450" r="1485900" b="914400" />',
  );
assert.notEqual(literalTextRectangleXml, customGeometryXml);
const literalTextRectangleFile = await PresentationFile.patchPptx(customGeometryPptx, [{ path: "ppt/slides/slide1.xml", xml: literalTextRectangleXml }]);
const literalTextRectanglePresentation = await PresentationFile.importPptx(literalTextRectangleFile);
const importedLiteralTextRectangle = itemByName(literalTextRectanglePresentation.slides.getItem(0).shapes.items, "literal-custom-path");
assert.deepEqual(importedLiteralTextRectangle.textRectangle, { left: 24, top: 18, right: 156, bottom: 96 });
const preservedLiteralTextRectangle = await PresentationFile.exportPptx(literalTextRectanglePresentation);
const preservedLiteralTextRectangleZip = await JSZip.loadAsync(preservedLiteralTextRectangle.bytes);
assert.match(await preservedLiteralTextRectangleZip.file("ppt/slides/slide1.xml").async("text"), /<a:rect l="228600" t="171450" r="1485900" b="914400"\s*\/>/);
importedLiteralTextRectangle.textRectangle.right = 157;
const editedLiteralTextRectangle = await PresentationFile.exportPptx(literalTextRectanglePresentation);
const editedLiteralTextRectangleZip = await JSZip.loadAsync(editedLiteralTextRectangle.bytes);
assert.match(await editedLiteralTextRectangleZip.file("ppt/slides/slide1.xml").async("text"), /<a:gd name="officeKitTextRight" fmla="\*\/ 1495425 w 1714500"\s*\/>/);
assert.deepEqual((await PresentationFile.importPptx(editedLiteralTextRectangle)).slides.getItem(0).shapes.items[0].textRectangle, { left: 24, top: 18, right: 157, bottom: 96 });
const formulaTextRectangleXml = customGeometryXml.replace('fmla="*/ 228600 w 1714500"', 'fmla="wd4"');
assert.notEqual(formulaTextRectangleXml, customGeometryXml);
const formulaTextRectangleFile = await PresentationFile.patchPptx(customGeometryPptx, [{ path: "ppt/slides/slide1.xml", xml: formulaTextRectangleXml }]);
const formulaTextRectanglePresentation = await PresentationFile.importPptx(formulaTextRectangleFile);
const opaqueFormulaTextRectangle = itemByName(formulaTextRectanglePresentation.slides.getItem(0).shapes.items, "literal-custom-path");
assert.equal(opaqueFormulaTextRectangle.customPaths.length, 0);
assert.equal(opaqueFormulaTextRectangle.textRectangle, undefined);
const preservedFormulaTextRectangle = await PresentationFile.exportPptx(formulaTextRectanglePresentation);
const preservedFormulaTextRectangleZip = await JSZip.loadAsync(preservedFormulaTextRectangle.bytes);
assert.match(await preservedFormulaTextRectangleZip.file("ppt/slides/slide1.xml").async("text"), /<a:gd name="officeKitTextLeft" fmla="wd4"\s*\/>/);
assert.match(await preservedFormulaTextRectangleZip.file("ppt/slides/slide1.xml").async("text"), /<a:rect l="officeKitTextLeft" t="officeKitTextTop" r="officeKitTextRight" b="officeKitTextBottom"\s*\/>/);
opaqueFormulaTextRectangle.textRectangle = { left: 24, top: 18, right: 156, bottom: 96 };
await assert.rejects(
  () => PresentationFile.exportPptx(formulaTextRectanglePresentation),
  (error) => error?.code === "unsupported_presentation_edit",
);
const mismatchedGuideTextRectangleXml = customGeometryXml.replace('fmla="*/ 228600 w 1714500"', 'fmla="*/ 228600 w 1714499"');
assert.notEqual(mismatchedGuideTextRectangleXml, customGeometryXml);
const mismatchedGuideTextRectangleFile = await PresentationFile.patchPptx(customGeometryPptx, [{ path: "ppt/slides/slide1.xml", xml: mismatchedGuideTextRectangleXml }]);
const mismatchedGuideTextRectanglePresentation = await PresentationFile.importPptx(mismatchedGuideTextRectangleFile);
const opaqueMismatchedGuideTextRectangle = itemByName(mismatchedGuideTextRectanglePresentation.slides.getItem(0).shapes.items, "literal-custom-path");
assert.equal(opaqueMismatchedGuideTextRectangle.customPaths.length, 0);
assert.equal(opaqueMismatchedGuideTextRectangle.textRectangle, undefined);
const preservedMismatchedGuideTextRectangle = await PresentationFile.exportPptx(mismatchedGuideTextRectanglePresentation);
const preservedMismatchedGuideTextRectangleZip = await JSZip.loadAsync(preservedMismatchedGuideTextRectangle.bytes);
assert.match(await preservedMismatchedGuideTextRectangleZip.file("ppt/slides/slide1.xml").async("text"), /<a:gd name="officeKitTextLeft" fmla="\*\/ 228600 w 1714499"\s*\/>/);
const childBearingTextRectangleXml = customGeometryXml.replace(
  '<a:rect l="officeKitTextLeft" t="officeKitTextTop" r="officeKitTextRight" b="officeKitTextBottom" />',
  '<a:rect l="officeKitTextLeft" t="officeKitTextTop" r="officeKitTextRight" b="officeKitTextBottom"><a:extLst /></a:rect>',
);
assert.notEqual(childBearingTextRectangleXml, customGeometryXml);
const childBearingTextRectangleFile = await PresentationFile.patchPptx(customGeometryPptx, [{ path: "ppt/slides/slide1.xml", xml: childBearingTextRectangleXml }]);
const childBearingTextRectanglePresentation = await PresentationFile.importPptx(childBearingTextRectangleFile);
const opaqueChildBearingTextRectangle = itemByName(childBearingTextRectanglePresentation.slides.getItem(0).shapes.items, "literal-custom-path");
assert.equal(opaqueChildBearingTextRectangle.customPaths.length, 0);
assert.equal(opaqueChildBearingTextRectangle.textRectangle, undefined);
const preservedChildBearingTextRectangle = await PresentationFile.exportPptx(childBearingTextRectanglePresentation);
const preservedChildBearingTextRectangleZip = await JSZip.loadAsync(preservedChildBearingTextRectangle.bytes);
assert.match(await preservedChildBearingTextRectangleZip.file("ppt/slides/slide1.xml").async("text"), /<a:rect l="officeKitTextLeft" t="officeKitTextTop" r="officeKitTextRight" b="officeKitTextBottom"><a:extLst\s*\/><\/a:rect>/);
const extraAttributeTextRectangleXml = customGeometryXml.replace('b="officeKitTextBottom"', 'b="officeKitTextBottom" data="unexpected"');
assert.notEqual(extraAttributeTextRectangleXml, customGeometryXml);
const extraAttributeTextRectangleFile = await PresentationFile.patchPptx(customGeometryPptx, [{ path: "ppt/slides/slide1.xml", xml: extraAttributeTextRectangleXml }]);
const extraAttributeTextRectanglePresentation = await PresentationFile.importPptx(extraAttributeTextRectangleFile);
const opaqueExtraAttributeTextRectangle = itemByName(extraAttributeTextRectanglePresentation.slides.getItem(0).shapes.items, "literal-custom-path");
assert.equal(opaqueExtraAttributeTextRectangle.customPaths.length, 0);
assert.equal(opaqueExtraAttributeTextRectangle.textRectangle, undefined);
const preservedExtraAttributeTextRectangle = await PresentationFile.exportPptx(extraAttributeTextRectanglePresentation);
const preservedExtraAttributeTextRectangleZip = await JSZip.loadAsync(preservedExtraAttributeTextRectangle.bytes);
assert.match(await preservedExtraAttributeTextRectangleZip.file("ppt/slides/slide1.xml").async("text"), /<a:rect\b[^>]*\bdata="unexpected"/);
const missingEdgeTextRectangleXml = customGeometryXml.replace(' b="officeKitTextBottom"', "");
assert.notEqual(missingEdgeTextRectangleXml, customGeometryXml);
const missingEdgeTextRectangleFile = await PresentationFile.patchPptx(customGeometryPptx, [{ path: "ppt/slides/slide1.xml", xml: missingEdgeTextRectangleXml }]);
const missingEdgeTextRectanglePresentation = await PresentationFile.importPptx(missingEdgeTextRectangleFile);
const opaqueMissingEdgeTextRectangle = itemByName(missingEdgeTextRectanglePresentation.slides.getItem(0).shapes.items, "literal-custom-path");
assert.equal(opaqueMissingEdgeTextRectangle.customPaths.length, 0);
assert.equal(opaqueMissingEdgeTextRectangle.textRectangle, undefined);
const preservedMissingEdgeTextRectangle = await PresentationFile.exportPptx(missingEdgeTextRectanglePresentation);
const preservedMissingEdgeTextRectangleZip = await JSZip.loadAsync(preservedMissingEdgeTextRectangle.bytes);
assert.match(await preservedMissingEdgeTextRectangleZip.file("ppt/slides/slide1.xml").async("text"), /<a:rect l="officeKitTextLeft" t="officeKitTextTop" r="officeKitTextRight"\s*\/>/);
const childBearingArcXml = customGeometryXml.replace(
  '<a:arcTo wR="3000" hR="4000" stAng="5400000" swAng="21600000" />',
  '<a:arcTo wR="3000" hR="4000" stAng="5400000" swAng="21600000"><a:extLst /></a:arcTo>',
);
assert.notEqual(childBearingArcXml, customGeometryXml);
const childBearingArcFile = await PresentationFile.patchPptx(customGeometryPptx, [{ path: "ppt/slides/slide1.xml", xml: childBearingArcXml }]);
const childBearingArcGeometry = await PresentationFile.importPptx(childBearingArcFile);
const opaqueChildBearingArc = itemByName(childBearingArcGeometry.slides.getItem(0).shapes.items, "literal-custom-path");
assert.equal(opaqueChildBearingArc.customPaths.length, 0);
const preservedChildBearingArc = await PresentationFile.exportPptx(childBearingArcGeometry);
const preservedChildBearingArcZip = await JSZip.loadAsync(preservedChildBearingArc.bytes);
assert.match(await preservedChildBearingArcZip.file("ppt/slides/slide1.xml").async("text"), /<a:arcTo wR="3000" hR="4000" stAng="5400000" swAng="21600000"><a:extLst\s*\/><\/a:arcTo>/);
const mixedQuadraticChildXml = customGeometryXml.replace('<a:pt x="18000" y="12000" />', '<a:extLst />');
assert.notEqual(mixedQuadraticChildXml, customGeometryXml);
const mixedQuadraticChildFile = await PresentationFile.patchPptx(customGeometryPptx, [{ path: "ppt/slides/slide1.xml", xml: mixedQuadraticChildXml }]);
const mixedQuadraticChildGeometry = await PresentationFile.importPptx(mixedQuadraticChildFile);
const opaqueMixedQuadraticChild = itemByName(mixedQuadraticChildGeometry.slides.getItem(0).shapes.items, "literal-custom-path");
assert.equal(opaqueMixedQuadraticChild.customPaths.length, 0);
const preservedMixedQuadraticChild = await PresentationFile.exportPptx(mixedQuadraticChildGeometry);
const preservedMixedQuadraticChildZip = await JSZip.loadAsync(preservedMixedQuadraticChild.bytes);
assert.match(await preservedMixedQuadraticChildZip.file("ppt/slides/slide1.xml").async("text"), /<a:quadBezTo><a:pt x="21000" y="6000"\s*\/><a:extLst\s*\/><\/a:quadBezTo>/);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "custom", customPaths: [{ width: 100, height: 100, commands: [{ arcTo: {} }] }] }),
  /arcTo\.widthRadius must be a safe integer/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "custom", customPaths: [{ width: 100, height: 100, commands: [{ arcTo: { widthRadius: 10, heightRadius: 10, startAngle: 0, sweepAngle: 5_400_000 } }] }] }),
  /arcTo requires an established current point/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "custom", customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 10, y: 10 } }, { arcTo: { widthRadius: 0, heightRadius: 10, startAngle: 0, sweepAngle: 5_400_000 } }] }] }),
  /arcTo radii must evaluate to positive values/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "custom", customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 10, y: 10 } }, { arcTo: { widthRadius: 10, heightRadius: 10, startAngle: 0, sweepAngle: 21_600_001 } }] }] }),
  /no greater than one full DrawingML turn/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "custom", customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 10, y: 10 } }, { arcTo: { widthRadius: 10, heightRadius: 10, startAngle: 0, sweepAngle: 0 } }] }] }),
  /sweepAngle must be non-zero/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "custom", customPaths: [{ width: 100, height: 100, commands: [{ toString: {} }] }] }),
  /unsupported command toString/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "custom", customPaths: [{ width: 100, height: 100, commands: [{ quadraticBezTo: { y1: 10, x: 20, y: 30 } }] }] }),
  /quadraticBezTo\.x1 must be a safe integer/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "custom", customPaths: [{ width: -1, height: 100, commands: [{ close: true }] }] }),
  /width must be positive when supplied/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "custom", customPaths: [{ width: 100, height: 100, fillMode: "lighten", commands: [{ close: true }] }] }),
  /fillMode must be normal or none/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "custom", customPaths: [{ width: 100, height: 100, stroke: "false", commands: [{ close: true }] }] }),
  /stroke must be a boolean/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "custom", customPaths: [{ width: 100, height: 100, extrusionAllowed: 1, commands: [{ close: true }] }] }),
  /extrusionAllowed must be a boolean/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "rect", textRectangle: { left: 0, top: 0, right: 10, bottom: 10 } }),
  /only for custom geometry shapes/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "rect", customConnectionSites: [{ angle: 0, x: 10, y: 10 }] }),
  /only for custom geometry shapes/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "rect", customAdjustmentHandles: [{ kind: "xy", xAdjustment: "adj", x: 10, y: 10 }] }),
  /declared custom adjustment/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    position: { left: 0, top: 0, width: 100, height: 100 },
    customAdjustments: [{ name: "adj", formula: "val 50" }],
    customAdjustmentHandles: Array.from({ length: 1_025 }, () => ({ kind: "xy", xAdjustment: "adj", x: 10, y: 10 })),
    customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 0, y: 0 } }] }],
  }),
  /at most 1024 entries/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    position: { left: 0, top: 0, width: 100, height: 100 },
    customAdjustments: [{ name: "adj", formula: "val 50" }],
    customAdjustmentHandles: [{ kind: "xy", xAdjustment: "missing", minX: 0, maxX: 100, x: 10, y: 10 }],
    customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 0, y: 0 } }] }],
  }),
  /declared custom adjustment/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    position: { left: 0, top: 0, width: 100, height: 100 },
    customAdjustments: [{ name: "adj", formula: "val 50" }],
    customAdjustmentHandles: [{ kind: "xy", xAdjustment: "adj", minX: 0, x: 10, y: 10 }],
    customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 0, y: 0 } }] }],
  }),
  /minX and maxX must be supplied together/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    position: { left: 0, top: 0, width: 100, height: 100 },
    customAdjustments: [{ name: "adj", formula: "val 150" }],
    customAdjustmentHandles: [{ kind: "xy", xAdjustment: "adj", minX: 0, maxX: 100, x: 10, y: 10 }],
    customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 0, y: 0 } }] }],
  }),
  /must evaluate inside its minX\/maxX range/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    position: { left: 0, top: 0, width: 100, height: 100 },
    customAdjustments: [{ name: "adj", formula: "val 50" }],
    customAdjustmentHandles: [{ kind: "xy", xAdjustment: "adj", minX: 0, maxX: 100, x: 101, y: 10 }],
    customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 0, y: 0 } }] }],
  }),
  /inside the custom shape frame/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    position: { left: 0, top: 0, width: 100, height: 100 },
    customAdjustments: [{ name: "adj", formula: "val 50" }],
    customAdjustmentHandles: [{ kind: "polar", radialAdjustment: "adj", minRadius: -1, maxRadius: 100, x: 10, y: 10 }],
    customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 0, y: 0 } }] }],
  }),
  /must evaluate to non-negative values/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    position: { left: 0, top: 0, width: 100, height: 100 },
    customAdjustments: [{ name: "adj", formula: "val -1" }],
    customAdjustmentHandles: [{ kind: "polar", radialAdjustment: "adj", x: 10, y: 10 }],
    customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 0, y: 0 } }] }],
  }),
  /radialAdjustment must evaluate to a non-negative value/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    position: { left: 0, top: 0, width: 100, height: 100 },
    customAdjustments: [{ name: "adj", formula: "val 21600001" }],
    customAdjustmentHandles: [{ kind: "polar", angleAdjustment: "adj", x: 10, y: 10 }],
    customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 0, y: 0 } }] }],
  }),
  /angleAdjustment must evaluate within one full DrawingML turn/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    position: { left: 0, top: 0, width: 100, height: 100 },
    customConnectionSites: Array.from({ length: 1_025 }, () => ({ angle: 0, x: 10, y: 10 })),
    customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 0, y: 0 } }] }],
  }),
  /at most 1024 entries/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    position: { left: 0, top: 0, width: 100, height: 100 },
    customConnectionSites: [{ angle: 361, x: 10, y: 10 }],
    customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 0, y: 0 } }] }],
  }),
  /degree value from -360 through 360/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    position: { left: 0, top: 0, width: 100, height: 100 },
    customConnectionSites: [{ angle: 0, x: 101, y: 10 }],
    customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 0, y: 0 } }] }],
  }),
  /inside the custom shape frame/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    position: { left: 0, top: 0, width: 100, height: 100 },
    customConnectionSites: [{ angle: 0, x: "missingGuide", y: 10 }],
    customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 0, y: 0 } }] }],
  }),
  /DrawingML built-in or declared guide reference/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "custom", textRectangle: { left: 0, top: 0, right: 10, bottom: 10, width: 10 }, customPaths: [{ width: 100, height: 100, commands: [{ close: true }] }] }),
  /textRectangle has unsupported fields: width/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "custom", textRectangle: { left: 10, top: 0, right: 10, bottom: 10 }, customPaths: [{ width: 100, height: 100, commands: [{ close: true }] }] }),
  /right must be greater than left/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    position: { left: 0, top: 0, width: 100, height: 100 },
    textRectangle: { left: "missingRectGuide", top: "t", right: "r", bottom: "b" },
    customPaths: [{ width: 100, height: 100, commands: [{ close: true }] }],
  }),
  /DrawingML built-in or declared guide reference/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    position: { left: 0, top: 0, width: 100, height: 100 },
    customGuides: [{ name: "near", formula: "val 80" }, { name: "far", formula: "val 20" }],
    textRectangle: { left: "near", top: "t", right: "far", bottom: "b" },
    customPaths: [{ width: 100, height: 100, commands: [{ close: true }] }],
  }),
  /right must be greater than left/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    customGuides: [{ name: "first", formula: "val later" }, { name: "later", formula: "val 1" }],
    customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 0, y: 0 } }] }],
  }),
  /unknown or forward guide later/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    customGuides: [{ name: "officeKitCollision", formula: "val 1" }],
    customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 0, y: 0 } }] }],
  }),
  /reserved officeKit prefix/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: "missingGuide", y: 0 } }] }],
  }),
  /must be a number or a DrawingML built-in or declared guide reference/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    customGuides: [{ name: "bad", formula: "*/ w 1 0" }],
    customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 0, y: 0 } }] }],
  }),
  /divides by zero/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "custom",
    customGuides: [{ name: "bad", formula: "max w" }],
    customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 0, y: 0 } }] }],
  }),
  /operator max requires 2 operands/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({
    geometry: "rect",
    customAdjustments: [{ name: "adj", formula: "val 1" }],
  }),
  /available only for custom geometry shapes/,
);
const originalFormulaGuide = formulaGeometryShape.customGuides[0].formula;
formulaGeometryShape.customGuides[0].formula = "*/ w 1 0";
assert.throws(() => formulaGeometryShape.inspectRecord(), /divides by zero/);
await assert.rejects(() => PresentationFile.exportPptx(formulaGeometryPresentation), /divides by zero/);
formulaGeometryShape.customGuides[0].formula = originalFormulaGuide;
const mutatedArcPresentation = Presentation.create({ slideSize: { width: 200, height: 120 } });
const mutatedArcSlide = mutatedArcPresentation.slides.add({ name: "Mutated arc" });
const mutatedArcShape = mutatedArcSlide.shapes.add({
  geometry: "custom",
  position: { left: 10, top: 10, width: 100, height: 80 },
  customPaths: [{
    width: 100,
    height: 100,
    commands: [
      { moveTo: { x: 50, y: 90 } },
      { arcTo: { widthRadius: 30, heightRadius: 40, startAngle: 5_400_000, sweepAngle: 10_800_000 } },
    ],
  }],
});
mutatedArcShape.customPaths[0].commands.shift();
await assert.rejects(() => mutatedArcSlide.export(), /arcTo requires an established current point/);
await assert.rejects(() => PresentationFile.exportPptx(mutatedArcPresentation), /arcTo requires an established current point/);
const mutatedPaintPresentation = Presentation.create({ slideSize: { width: 200, height: 120 } });
const mutatedPaintSlide = mutatedPaintPresentation.slides.add({ name: "Mutated path paint" });
const mutatedPaintShape = mutatedPaintSlide.shapes.add({
  geometry: "custom",
  position: { left: 10, top: 10, width: 100, height: 80 },
  customPaths: [{
    width: 100,
    height: 100,
    fillMode: "none",
    commands: [{ moveTo: { x: 10, y: 10 } }, { lineTo: { x: 90, y: 90 } }],
  }],
});
mutatedPaintShape.customPaths[0].stroke = "false";
await assert.rejects(() => mutatedPaintSlide.export(), /stroke must be a boolean/);
await assert.rejects(() => PresentationFile.exportPptx(mutatedPaintPresentation), /stroke must be a boolean/);
const mutatedTextRectanglePresentation = Presentation.create({ slideSize: { width: 200, height: 120 } });
const mutatedTextRectangleSlide = mutatedTextRectanglePresentation.slides.add({ name: "Mutated text rectangle" });
const mutatedTextRectangleShape = mutatedTextRectangleSlide.shapes.add({
  geometry: "custom",
  position: { left: 10, top: 10, width: 100, height: 80 },
  textRectangle: { left: 10, top: 10, right: 90, bottom: 70 },
  customPaths: [{ width: 100, height: 100, commands: [{ moveTo: { x: 10, y: 10 } }, { lineTo: { x: 90, y: 90 } }] }],
});
mutatedTextRectangleShape.textRectangle.right = 10;
assert.throws(() => mutatedTextRectangleShape.inspectRecord(), /right must be greater than left/);
assert.throws(() => mutatedTextRectangleShape.layoutJson(), /right must be greater than left/);
await assert.rejects(() => mutatedTextRectangleSlide.export(), /right must be greater than left/);
await assert.rejects(() => PresentationFile.exportPptx(mutatedTextRectanglePresentation), /right must be greater than left/);

// Groups are a recursive DrawingML ownership boundary, not flattened children
// with synthetic parent IDs. The public model keeps child coordinates local and
// OfficeKit authors/imports native p:grpSp trees with fixed-topology edits.
const groupedPresentation = Presentation.create({ slideSize: { width: 960, height: 540 } });
const groupedSlide = groupedPresentation.slides.add({ name: "Native group tree" });
const authoredGroup = groupedSlide.groups.add({
  name: "Agent evidence group",
  position: { left: 100, top: 80, width: 600, height: 320 },
  childFrame: { left: -100, top: 50, width: 1200, height: 640 },
  accessibility: { title: "Agent evidence flow", description: "Grouped visual containing before, target, evidence, table, and chart objects." },
});
const groupedBefore = authoredGroup.shapes.add({
  name: "grouped-before",
  geometry: "roundRect",
  position: { left: 0, top: 100, width: 300, height: 120 },
  fill: "#DBEAFE",
  line: { fill: "#2563EB", width: 2 },
  text: "Before",
});
const groupedTarget = authoredGroup.shapes.add({
  name: "grouped-target",
  geometry: "rect",
  position: { left: 450, top: 100, width: 300, height: 120 },
  fill: "#DCFCE7",
  line: { fill: "#16A34A", width: 2 },
  text: "Target",
});
const groupedConnector = authoredGroup.connectors.add({
  name: "grouped-connector",
  connectorType: "straight",
  from: groupedBefore,
  to: groupedTarget,
  start: { x: 300, y: 160 },
  end: { x: 450, y: 160 },
  line: { fill: "#334155", width: 2, endArrow: "triangle" },
  accessibility: { title: "Before-to-target direction", description: "Arrow connecting the before state to the target state." },
});
authoredGroup.images.add({
  name: "grouped-image",
  alt: "Grouped image evidence",
  position: { left: 800, top: 100, width: 120, height: 120 },
  fit: "stretch",
  dataUrl: PNG,
});
authoredGroup.tables.add({
  name: "grouped-table",
  position: { left: 0, top: 300, width: 400, height: 180 },
  values: [["Gate", "State"], ["Import", "Before"]],
  styleOptions: { headerRow: true, bandedRows: true },
  accessibility: { title: "Grouped gate table", description: "Import gate state inside the evidence group." },
});
authoredGroup.charts.add("combo", {
  name: "grouped-chart",
  title: "Grouped readiness",
  accessibility: { title: "Grouped readiness chart", description: "Create and edit readiness scores inside the evidence group." },
  position: { left: 450, top: 300, width: 350, height: 200 },
  categories: ["Create", "Edit"],
  series: [
    { name: "Score", chartType: "bar", values: [7, 9], color: "#7C3AED" },
    {
      name: "Review", chartType: "line", axisGroup: "secondary", values: [5, 8],
      line: { fill: "#0F766E", width: 2 },
      marker: { symbol: "circle", size: 6, fill: "#0F766E" },
    },
  ],
  axes: {
    category: { title: "Stage" },
    value: { title: "Score" },
    secondary: { category: { title: "Stage" }, value: { title: "Review", min: 0, max: 10, majorUnit: 2 } },
  },
  legend: false,
});
const nestedGroup = authoredGroup.groups.add({
  name: "nested-group",
  position: { left: 850, top: 300, width: 250, height: 220 },
  childFrame: { left: 0, top: 0, width: 250, height: 220 },
  accessibility: { description: "Nested custom-shape evidence." },
});
const nestedCustomShape = nestedGroup.shapes.add({
  name: "nested-shape",
  geometry: "custom",
  position: { left: 20, top: 30, width: 200, height: 120 },
  fill: "#FCE7F3",
  line: { fill: "#BE185D", width: 1 },
  text: "Nested",
  textRectangle: { left: 30, top: 20, right: 170, bottom: 100 },
  customPaths: [{
    width: 100,
    height: 100,
    commands: [
      { moveTo: { x: 50, y: 0 } },
      { lineTo: { x: 100, y: 50 } },
      { lineTo: { x: 50, y: 100 } },
      { lineTo: { x: 0, y: 50 } },
      { close: {} },
    ],
  }],
});
assert.equal(groupedPresentation.resolve(groupedBefore.id), groupedBefore);
assert.match(groupedPresentation.inspect({ kind: "groupShape,shape,connector,table,chart,image", maxChars: 20_000 }).ndjson, /Agent evidence group/);
assert.match(authoredGroup.toSvg(), /translate\(100 80\) scale\(0\.5 0\.5\) translate\(100 -50\)/);
nestedCustomShape.textRectangle.bottom = 21;
const groupedCustomTextOverflow = groupedPresentation.validateLayout().issues.find((issue) => issue.id === nestedCustomShape.id && issue.type === "textOverflow");
assert.deepEqual(groupedCustomTextOverflow?.bbox, [600, 230, 70, 0.5]);
nestedCustomShape.textRectangle.bottom = 100;
const groupedVerification = groupedPresentation.verify();
assert.equal(groupedVerification.ok, true, JSON.stringify(groupedVerification.issues));
const groupedLayoutValidation = groupedPresentation.validateLayout();
assert.equal(groupedLayoutValidation.ok, true, JSON.stringify(groupedLayoutValidation.issues));

const groupedFirstExport = await PresentationFile.exportPptx(groupedPresentation);
const groupedFirstZip = await JSZip.loadAsync(new Uint8Array(await groupedFirstExport.arrayBuffer()));
const groupedFirstXml = await groupedFirstZip.file("ppt/slides/slide1.xml").async("text");
assert.equal((groupedFirstXml.match(/<p:grpSp>/g) || []).length, 2);
assert.match(groupedFirstXml, /<a:chOff x="-952500" y="476250"\s*\/>/);
assert.match(groupedFirstXml, /<a:chExt cx="11430000" cy="6096000"\s*\/>/);
assert.match(groupedFirstXml, /<p:cNvPr\b(?=[^>]*\bname="grouped-table")(?=[^>]*\btitle="Grouped gate table")/);
assert.match(groupedFirstXml, /<p:cNvPr\b(?=[^>]*\bname="grouped-chart")(?=[^>]*\btitle="Grouped readiness chart")/);
assert.match(groupedFirstXml, /<p:cNvPr\b(?=[^>]*\bname="Agent evidence group")(?=[^>]*\btitle="Agent evidence flow")(?=[^>]*\bdescr="Grouped visual containing before, target, evidence, table, and chart objects\.")/);
assert.match(groupedFirstXml, /<p:cNvPr\b(?=[^>]*\bname="grouped-connector")(?=[^>]*\btitle="Before-to-target direction")(?=[^>]*\bdescr="Arrow connecting the before state to the target state\.")/);
assert.match(groupedFirstXml, /<p:cNvPr\b(?=[^>]*\bname="nested-group")(?=[^>]*\bdescr="Nested custom-shape evidence\.")/);

let groupedImported = await PresentationFile.importPptx(groupedFirstExport);
let importedGroup = itemByName(groupedImported.slides.getItem(0).groups.items, "Agent evidence group");
assert.deepEqual(importedGroup.children.map((child) => child.layoutJson().kind), ["textbox", "textbox", "connector", "image", "table", "chart", "groupShape"]);
assert.deepEqual(importedGroup.childFrame, { left: -100, top: 50, width: 1200, height: 640 });
assert.equal(groupedImported.resolve(itemByName(importedGroup.shapes.items, "grouped-before").id).text.value, "Before");
assert.deepEqual(importedGroup.accessibilityCapability, { sourceBound: true, editable: true, addable: true });
assert.deepEqual(itemByName(importedGroup.connectors.items, "grouped-connector").accessibilityCapability, { sourceBound: true, editable: true, addable: true });
assert.deepEqual(itemByName(importedGroup.groups.items, "nested-group").accessibilityCapability, { sourceBound: true, editable: true, addable: true });
assert.deepEqual(importedGroup.accessibility, authoredGroup.accessibility);
assert.deepEqual(itemByName(importedGroup.connectors.items, "grouped-connector").accessibility, groupedConnector.accessibility);
assert.deepEqual(itemByName(importedGroup.groups.items, "nested-group").accessibility, nestedGroup.accessibility);
const groupedNoOp = await PresentationFile.exportPptx(groupedImported);
assert.deepEqual(groupedNoOp.bytes, groupedFirstExport.bytes, "unchanged imported group accessibility metadata must return the exact source package");
const groupedNativeLeafImported = await PresentationFile.importPptx(groupedFirstExport);
const groupedNativeGroup = itemByName(groupedNativeLeafImported.slides.getItem(0).groups.items, "Agent evidence group");
const groupedNativeBefore = itemByName(groupedNativeGroup.shapes.items, "grouped-before");
const groupedNativeLeafRecords = groupedNativeLeafImported.inspect({ includeNativeLeaves: true, target: groupedNativeBefore.id }).ndjson
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line))
  .filter((record) => record.kind === "nativeLeaf");
const groupedBeforeLeaf = groupedNativeLeafRecords.find((record) => record.value === "Before");
assert.ok(groupedBeforeLeaf);
assert.equal(groupedBeforeLeaf.parentGroupId, groupedNativeGroup.id);
groupedNativeLeafImported.editNativeLeaf(groupedBeforeLeaf.targetId, groupedBeforeLeaf.leafId, {
  expectedHash: groupedBeforeLeaf.expectedHash,
  value: "After native leaf",
});
const groupedNativeLeafOutput = await PresentationFile.exportPptx(groupedNativeLeafImported);
assert.deepEqual(groupedNativeLeafOutput.metadata.editPlan.operations[0].shapeTreePath.length, 2);
assert.deepEqual(groupedNativeLeafOutput.metadata.editPlan.operations[0].footprint.shapeTreePath, groupedNativeLeafOutput.metadata.editPlan.operations[0].shapeTreePath);
const groupedNativeLeafXml = await (await JSZip.loadAsync(groupedNativeLeafOutput.bytes)).file("ppt/slides/slide1.xml").async("text");
assert.equal(groupedNativeLeafXml.replace("After native leaf", "Before"), groupedFirstXml);
const groupedNativeLeafRoundTrip = await PresentationFile.importPptx(groupedNativeLeafOutput);
assert.equal(groupedNativeLeafRoundTrip.resolve(groupedBeforeLeaf.targetId).text.value, "After native leaf");

const groupedColorLeafImported = await PresentationFile.importPptx(groupedFirstExport);
const groupedColorGroup = itemByName(groupedColorLeafImported.slides.getItem(0).groups.items, "Agent evidence group");
const groupedColorShape = itemByName(groupedColorGroup.shapes.items, "grouped-before");
const groupedScalarLeaves = groupedColorLeafImported.inspect({ includeNativeLeaves: true, target: groupedColorShape.id }).ndjson
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line))
  .filter((record) => record.kind === "nativeLeaf");
assert.deepEqual(new Set(groupedScalarLeaves.map((record) => record.leafKind)), new Set(["text", "fillRgb", "lineRgb"]));
const groupedFillLeaf = groupedScalarLeaves.find((record) => record.leafKind === "fillRgb");
assert.ok(groupedFillLeaf);
assert.equal(groupedFillLeaf.value, "#dbeafe");
assert.throws(
  () => groupedColorLeafImported.editNativeLeaf(groupedFillLeaf.targetId, groupedFillLeaf.leafId, { expectedHash: groupedFillLeaf.expectedHash, value: "#DBEAFE" }),
  (error) => error?.code === "presentation_native_leaf_noop",
);
assert.throws(
  () => groupedColorLeafImported.editNativeLeaf(groupedFillLeaf.targetId, groupedFillLeaf.leafId, { expectedHash: "0".repeat(64), value: "#A1B2C3" }),
  (error) => error?.code === "presentation_native_leaf_stale",
);
assert.throws(
  () => groupedColorLeafImported.editNativeLeaf(groupedFillLeaf.targetId, groupedFillLeaf.leafId, { expectedHash: groupedFillLeaf.expectedHash, value: "theme-accent" }),
  (error) => error?.code === "invalid_presentation_native_leaf_edit",
);
const groupedColorEdit = groupedColorLeafImported.editNativeLeaf(groupedFillLeaf.targetId, groupedFillLeaf.leafId, {
  expectedHash: groupedFillLeaf.expectedHash,
  value: "#A1B2C3",
});
assert.equal(groupedColorEdit.leafKind, "fillRgb");
const groupedColorOutput = await PresentationFile.exportPptx(groupedColorLeafImported);
const groupedColorOperation = groupedColorOutput.metadata.editPlan.operations[0];
assert.equal(groupedColorOperation.leafKind, "fillRgb");
assert.equal(groupedColorOperation.footprint.leafKind, "fillRgb");
assert.equal(groupedColorOperation.shapeTreePath.length, 2);
await assertOnlyDeclaredPptxFootprintChanged(groupedFirstExport, groupedColorOutput, groupedColorOperation);
const groupedColorRoundTrip = await PresentationFile.importPptx(groupedColorOutput);
assert.equal(itemByName(itemByName(groupedColorRoundTrip.slides.getItem(0).groups.items, "Agent evidence group").shapes.items, "grouped-before").fill.toLowerCase(), "#a1b2c3");
assert.throws(
  () => groupedColorRoundTrip.editNativeLeaf(groupedFillLeaf.targetId, groupedFillLeaf.leafId, { expectedHash: groupedFillLeaf.expectedHash, value: "#C3B2A1" }),
  (error) => error?.code === "presentation_native_leaf_not_issued",
);

const groupedDependentEditImported = await PresentationFile.importPptx(groupedFirstExport);
const groupedDependentGroup = itemByName(groupedDependentEditImported.slides.getItem(0).groups.items, "Agent evidence group");
const groupedDependentShape = itemByName(groupedDependentGroup.shapes.items, "grouped-before");
const groupedDependentFillLeaf = groupedDependentEditImported.inspect({ includeNativeLeaves: true, target: groupedDependentShape.id }).ndjson
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line))
  .find((record) => record.kind === "nativeLeaf" && record.leafKind === "fillRgb");
assert.ok(groupedDependentFillLeaf);
groupedDependentShape.position.top += 1;
assert.throws(
  () => groupedDependentEditImported.editNativeLeaf(groupedDependentFillLeaf.targetId, groupedDependentFillLeaf.leafId, {
    expectedHash: groupedDependentFillLeaf.expectedHash,
    value: "#ABCDEF",
  }),
  (error) => error?.code === "presentation_native_leaf_concurrent_change",
);
const groupedPostIssueImported = await PresentationFile.importPptx(groupedFirstExport);
const groupedPostIssueGroup = itemByName(groupedPostIssueImported.slides.getItem(0).groups.items, "Agent evidence group");
const groupedPostIssueShape = itemByName(groupedPostIssueGroup.shapes.items, "grouped-before");
const groupedPostIssueFillLeaf = groupedPostIssueImported.inspect({ includeNativeLeaves: true, target: groupedPostIssueShape.id }).ndjson
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line))
  .find((record) => record.kind === "nativeLeaf" && record.leafKind === "fillRgb");
assert.ok(groupedPostIssueFillLeaf);
groupedPostIssueImported.editNativeLeaf(groupedPostIssueFillLeaf.targetId, groupedPostIssueFillLeaf.leafId, {
  expectedHash: groupedPostIssueFillLeaf.expectedHash,
  value: "#ABCDEF",
});
groupedPostIssueShape.position.left += 1;
await assert.rejects(
  () => PresentationFile.exportPptx(groupedPostIssueImported),
  (error) => error?.code === "unsupported_presentation_native_leaf_edit",
  "an explicit native-leaf edit must never fall back to full presentation serialization when a dependent change escapes its Edit Plan",
);

const groupedGeometryLeafImported = await PresentationFile.importPptx(groupedFirstExport);
const groupedGeometryGroup = itemByName(groupedGeometryLeafImported.slides.getItem(0).groups.items, "Agent evidence group");
const groupedGeometryNestedGroup = itemByName(groupedGeometryGroup.groups.items, "nested-group");
const groupedGeometryShape = itemByName(groupedGeometryNestedGroup.shapes.items, "nested-shape");
const groupedGeometryLeaves = groupedGeometryLeafImported.inspect({ includeNativeLeaves: true, target: groupedGeometryShape.id }).ndjson
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line))
  .filter((record) => record.kind === "nativeLeaf");
const groupedLeftLeaf = groupedGeometryLeaves.find((record) => record.leafKind === "leftEmu");
const groupedWidthLeaf = groupedGeometryLeaves.find((record) => record.leafKind === "widthEmu");
assert.ok(groupedLeftLeaf);
assert.ok(groupedWidthLeaf);
assert.equal(groupedWidthLeaf.unit, "emu");
assert.throws(
  () => groupedGeometryLeafImported.editNativeLeaf(groupedWidthLeaf.targetId, groupedWidthLeaf.leafId, { expectedHash: groupedWidthLeaf.expectedHash, value: 0 }),
  (error) => error?.code === "invalid_presentation_native_leaf_edit",
);
const nextLeftEmu = groupedLeftLeaf.value + 9_525;
const groupedGeometryEdit = groupedGeometryLeafImported.editNativeLeaf(groupedLeftLeaf.targetId, groupedLeftLeaf.leafId, {
  expectedHash: groupedLeftLeaf.expectedHash,
  value: nextLeftEmu,
});
assert.equal(groupedGeometryEdit.value, nextLeftEmu);
assert.equal(groupedGeometryEdit.unit, "emu");
const groupedGeometryOutput = await PresentationFile.exportPptx(groupedGeometryLeafImported);
const groupedGeometryOperation = groupedGeometryOutput.metadata.editPlan.operations[0];
assert.equal(groupedGeometryOperation.leafKind, "leftEmu");
assert.equal(groupedGeometryOperation.footprint.leafKind, "leftEmu");
await assertOnlyDeclaredPptxFootprintChanged(groupedFirstExport, groupedGeometryOutput, groupedGeometryOperation);
const groupedGeometryRoundTrip = await PresentationFile.importPptx(groupedGeometryOutput);
assert.equal(itemByName(itemByName(itemByName(groupedGeometryRoundTrip.slides.getItem(0).groups.items, "Agent evidence group").groups.items, "nested-group").shapes.items, "nested-shape").position.left, nextLeftEmu / 9_525);
const groupedAccessibilitySourceSvg = await groupedImported.slides.getItem(0).export({ format: "svg" });
importedGroup.setAccessibilityMetadata({ title: "Reviewed agent evidence flow", description: null });
itemByName(importedGroup.connectors.items, "grouped-connector").setAccessibilityMetadata({ title: null, description: "Reviewed arrow from before to target." });
itemByName(importedGroup.groups.items, "nested-group").setAccessibilityMetadata({ title: "Nested evidence" });
const groupedAccessibilityOnlyExport = await PresentationFile.exportPptx(groupedImported);
groupedImported = await PresentationFile.importPptx(groupedAccessibilityOnlyExport);
const groupedAccessibilityOutputSvg = await groupedImported.slides.getItem(0).export({ format: "svg" });
assert.deepEqual(groupedAccessibilityOutputSvg.bytes, groupedAccessibilitySourceSvg.bytes, "group and connector accessibility edits must not alter model SVG output");
importedGroup = itemByName(groupedImported.slides.getItem(0).groups.items, "Agent evidence group");
assert.deepEqual(importedGroup.accessibility, { title: "Reviewed agent evidence flow" });
assert.deepEqual(itemByName(importedGroup.connectors.items, "grouped-connector").accessibility, { description: "Reviewed arrow from before to target." });
assert.deepEqual(itemByName(importedGroup.groups.items, "nested-group").accessibility, { title: "Nested evidence", description: "Nested custom-shape evidence." });

importedGroup.name = "Edited agent evidence group";
importedGroup.position.left = 120;
importedGroup.childFrame.left = -50;
itemByName(importedGroup.shapes.items, "grouped-before").text.set("After");
delete itemByName(importedGroup.connectors.items, "grouped-connector").line.endArrow;
itemByName(importedGroup.images.items, "grouped-image").alt = "Edited grouped image evidence";
const importedGroupedTable = itemByName(importedGroup.tables.items, "grouped-table");
assert.deepEqual(importedGroupedTable.accessibilityCapability, { sourceBound: true, editable: true, addable: true });
assert.equal(importedGroupedTable.accessibility.title, "Grouped gate table");
importedGroupedTable.cells.set(1, 1, "After");
importedGroupedTable.setAccessibilityMetadata({ description: "Edited import gate state inside the evidence group." });
const importedGroupedChart = itemByName(importedGroup.charts.items, "grouped-chart");
assert.equal(importedGroupedChart.chartType, "combo");
assert.deepEqual(importedGroupedChart.accessibilityCapability, { sourceBound: true, editable: true, addable: true });
assert.equal(importedGroupedChart.accessibility.title, "Grouped readiness chart");
assert.deepEqual(importedGroupedChart.series.map((series) => [series.chartType, series.axisGroup || "primary"]), [["bar", "primary"], ["line", "secondary"]]);
assert.equal(importedGroupedChart.axes.secondary.value.max, 10);
importedGroupedChart.setAccessibilityMetadata({ title: "Edited grouped readiness chart" });
importedGroupedChart.title = "Edited grouped readiness";
importedGroupedChart.series[0].values = [8, 10];
importedGroupedChart.series[1].values = [6, 9];
importedGroupedChart.axes.secondary.value.max = 12;
const importedNestedGroup = itemByName(importedGroup.groups.items, "nested-group");
assert.deepEqual(itemByName(importedNestedGroup.shapes.items, "nested-shape").textRectangle, { left: 30, top: 20, right: 170, bottom: 100 });
importedNestedGroup.position.top = 320;
const importedNestedShape = itemByName(importedNestedGroup.shapes.items, "nested-shape");
importedNestedShape.fill = "#FDE68A";
importedNestedShape.textRectangle.left = 40;

const groupedSecondExport = await PresentationFile.exportPptx(groupedImported);
const groupedRoundTrip = await PresentationFile.importPptx(groupedSecondExport);
const roundTripGroup = itemByName(groupedRoundTrip.slides.getItem(0).groups.items, "Edited agent evidence group");
assert.equal(roundTripGroup.position.left, 120);
assert.equal(roundTripGroup.childFrame.left, -50);
assert.deepEqual(roundTripGroup.accessibility, { title: "Reviewed agent evidence flow" });
assert.equal(itemByName(roundTripGroup.shapes.items, "grouped-before").text.value, "After");
assert.equal(itemByName(roundTripGroup.connectors.items, "grouped-connector").line.endArrow, undefined);
assert.deepEqual(itemByName(roundTripGroup.connectors.items, "grouped-connector").accessibility, { description: "Reviewed arrow from before to target." });
assert.equal(itemByName(roundTripGroup.images.items, "grouped-image").alt, "Edited grouped image evidence");
assert.equal(itemByName(roundTripGroup.tables.items, "grouped-table").values[1][1], "After");
assert.deepEqual(itemByName(roundTripGroup.tables.items, "grouped-table").accessibility, {
  title: "Grouped gate table",
  description: "Edited import gate state inside the evidence group.",
});
assert.equal(itemByName(roundTripGroup.charts.items, "grouped-chart").chartType, "combo");
assert.equal(itemByName(roundTripGroup.charts.items, "grouped-chart").accessibility.title, "Edited grouped readiness chart");
assert.deepEqual(itemByName(roundTripGroup.charts.items, "grouped-chart").series[0].values, [8, 10]);
assert.deepEqual(itemByName(roundTripGroup.charts.items, "grouped-chart").series[1].values, [6, 9]);
assert.equal(itemByName(roundTripGroup.charts.items, "grouped-chart").series[1].axisGroup, "secondary");
assert.equal(itemByName(roundTripGroup.charts.items, "grouped-chart").axes.secondary.value.max, 12);
assert.equal(itemByName(itemByName(roundTripGroup.groups.items, "nested-group").shapes.items, "nested-shape").fill, "#FDE68A");
assert.deepEqual(itemByName(roundTripGroup.groups.items, "nested-group").accessibility, { title: "Nested evidence", description: "Nested custom-shape evidence." });
assert.deepEqual(itemByName(itemByName(roundTripGroup.groups.items, "nested-group").shapes.items, "nested-shape").textRectangle, { left: 40, top: 20, right: 170, bottom: 100 });

const removedGroupedChild = roundTripGroup.children.pop();
await assert.rejects(
  () => PresentationFile.exportPptx(groupedRoundTrip),
  (error) => error?.code === "presentation_group_topology_changed",
);
roundTripGroup.children.push(removedGroupedChild);

// An irregular cNvPr leaf is source-owned independently of an otherwise
// canonical group or connector. Geometry/line edits preserve it byte-for-byte,
// while both the public setter and direct-state bypass fail closed.
const irregularGroupAccessibilityXml = groupedFirstXml
  .replace(/(<p:cNvPr\b[^>]*\bname="Agent evidence group")/, '$1 xmlns:fixture="urn:office-kit:group-connector-accessibility" fixture:group="kept"')
  .replace(/(<p:cNvPr\b[^>]*\bname="grouped-connector")/, '$1 xmlns:fixture="urn:office-kit:group-connector-accessibility" fixture:connector="kept"');
assert.notEqual(irregularGroupAccessibilityXml, groupedFirstXml);
const irregularGroupAccessibilityFile = await PresentationFile.patchPptx(groupedFirstExport, [{ path: "ppt/slides/slide1.xml", xml: irregularGroupAccessibilityXml }]);
const irregularGroupAccessibilityPresentation = await PresentationFile.importPptx(irregularGroupAccessibilityFile);
const irregularAccessibleGroup = itemByName(irregularGroupAccessibilityPresentation.slides.getItem(0).groups.items, "Agent evidence group");
const irregularAccessibleConnector = itemByName(irregularAccessibleGroup.connectors.items, "grouped-connector");
assert.equal(irregularAccessibleGroup.accessibility, undefined);
assert.equal(irregularAccessibleConnector.accessibility, undefined);
assert.deepEqual(irregularAccessibleGroup.accessibilityCapability, { sourceBound: true, editable: false, addable: false });
assert.deepEqual(irregularAccessibleConnector.accessibilityCapability, { sourceBound: true, editable: false, addable: false });
assert.throws(() => irregularAccessibleGroup.setAccessibilityMetadata({ title: "Do not flatten group metadata" }), /source-bound.*editable p:cNvPr profile/i);
assert.throws(() => irregularAccessibleConnector.setAccessibilityMetadata({ title: "Do not flatten connector metadata" }), /source-bound.*editable p:cNvPr profile/i);
irregularAccessibleGroup.position.left += 10;
irregularAccessibleConnector.line.width += 0.5;
const irregularGroupAccessibilityOtherEdit = await PresentationFile.exportPptx(irregularGroupAccessibilityPresentation);
const irregularGroupAccessibilityOtherXml = await (await JSZip.loadAsync(irregularGroupAccessibilityOtherEdit.bytes)).file("ppt/slides/slide1.xml").async("text");
assert.match(irregularGroupAccessibilityOtherXml, /fixture:group="kept"/);
assert.match(irregularGroupAccessibilityOtherXml, /fixture:connector="kept"/);
assert.match(irregularGroupAccessibilityOtherXml, /title="Agent evidence flow"/);
assert.match(irregularGroupAccessibilityOtherXml, /title="Before-to-target direction"/);
const irregularGroupAccessibilityBypass = await PresentationFile.importPptx(irregularGroupAccessibilityFile);
itemByName(irregularGroupAccessibilityBypass.slides.getItem(0).groups.items, "Agent evidence group").accessibility = { title: "Bypass group" };
await assert.rejects(() => PresentationFile.exportPptx(irregularGroupAccessibilityBypass), (error) => error?.code === "unsupported_presentation_edit");
const irregularConnectorAccessibilityBypass = await PresentationFile.importPptx(irregularGroupAccessibilityFile);
itemByName(itemByName(irregularConnectorAccessibilityBypass.slides.getItem(0).groups.items, "Agent evidence group").connectors.items, "grouped-connector").accessibility = { title: "Bypass connector" };
await assert.rejects(() => PresentationFile.exportPptx(irregularConnectorAccessibilityBypass), (error) => error?.code === "unsupported_presentation_edit");

const irregularGroupXml = groupedFirstXml.replace(
  /(<p:grpSp><p:nvGrpSpPr><p:cNvPr\b[^>]*name="Agent evidence group"[^>]*\/>[\s\S]*?<p:grpSpPr)(>)/,
  "$1 bwMode=\"gray\"$2",
);
assert.notEqual(irregularGroupXml, groupedFirstXml);
const irregularGroupFile = await PresentationFile.patchPptx(groupedFirstExport, [{ path: "ppt/slides/slide1.xml", xml: irregularGroupXml }]);
const irregularGroupZip = await JSZip.loadAsync(new Uint8Array(await irregularGroupFile.arrayBuffer()));
assert.match(await irregularGroupZip.file("ppt/slides/slide1.xml").async("text"), /<p:grpSpPr bwMode="gray">/);
const irregularGroupPresentation = await PresentationFile.importPptx(irregularGroupFile);
const irregularGroupSlide = irregularGroupPresentation.slides.getItem(0);
assert.equal(irregularGroupSlide.groups.items.length, 0);
const opaqueGroup = itemByName(irregularGroupSlide.nativeObjects.items, "Agent evidence group");
assert.equal(opaqueGroup.editable, false);
opaqueGroup.name = "Unsafe group edit";
await assert.rejects(
  () => PresentationFile.exportPptx(irregularGroupPresentation),
  (error) => error?.code === "unsupported_presentation_edit",
);

// Office 2019+ decorative classification is one presence-aware accessibility
// value across every modeled drawing object. Explicit false remains distinct
// from absence, and true cannot coexist with title/description.
const decorativePresentation = Presentation.create({ slideSize: { width: 960, height: 540 } });
const decorativeSlide = decorativePresentation.slides.add({ name: "Decorative object semantics" });
const classifiedShape = decorativeSlide.shapes.add({
  name: "classified-shape",
  position: { left: 40, top: 40, width: 180, height: 64 },
  text: "Meaningful status",
  accessibility: { title: "Meaningful status", decorative: false },
});
const decorativeImage = decorativeSlide.images.add({
  name: "decorative-image",
  position: { left: 250, top: 40, width: 96, height: 64 },
  dataUrl: PNG,
  fit: "stretch",
  prompt: "Generation prompt must not become alternative text",
  accessibility: { decorative: true },
});
const decorativeTable = decorativeSlide.tables.add({
  name: "decorative-table",
  position: { left: 380, top: 40, width: 160, height: 64 },
  values: [["Visual key"]],
  accessibility: { decorative: true },
});
const decorativeChart = decorativeSlide.charts.add("bar", {
  name: "decorative-chart",
  position: { left: 580, top: 40, width: 300, height: 180 },
  categories: ["Ready"],
  series: [{ name: "State", values: [1] }],
  legend: false,
  accessibility: { decorative: true },
});
const decorativeGroup = decorativeSlide.groups.add({
  name: "decorative-group",
  position: { left: 80, top: 280, width: 420, height: 140 },
  childFrame: { left: 0, top: 0, width: 420, height: 140 },
  accessibility: { decorative: true },
});
const decorativeFrom = decorativeGroup.shapes.add({
  name: "decorative-from",
  position: { left: 10, top: 40, width: 120, height: 60 },
  text: "A",
  accessibility: { title: "Workflow source node", decorative: false },
});
const decorativeTo = decorativeGroup.shapes.add({
  name: "decorative-to",
  position: { left: 290, top: 40, width: 120, height: 60 },
  text: "B",
  accessibility: { title: "Workflow destination node", decorative: false },
});
const decorativeConnector = decorativeGroup.connectors.add({
  name: "decorative-connector",
  connectorType: "straight",
  from: decorativeFrom,
  to: decorativeTo,
  start: { x: 130, y: 70 },
  end: { x: 290, y: 70 },
  line: { fill: "#94A3B8", width: 1 },
  accessibility: { decorative: true },
});
assert.equal(decorativeImage.alt, "", "a decorative image must not derive alternative text from its generation prompt");
assert.throws(
  () => classifiedShape.setAccessibilityMetadata({ decorative: true }),
  /cannot combine decorative: true with title or description/i,
);
assert.deepEqual(classifiedShape.accessibility, { title: "Meaningful status", decorative: false });
assert.throws(() => decorativeImage.setAccessibilityMetadata({ decorative: "true" }), /decorative must be a boolean/i);
const decorativeAudit = decorativePresentation.auditAccessibility({ maxChars: 20_000 });
assert.equal(decorativeAudit.machineCheckPassed, true);
assert.equal(decorativeAudit.conformanceClaimed, false);
assert.equal(decorativeAudit.manualReviewRequired, true);
assert.deepEqual(decorativeAudit.summary, {
  slides: 1,
  modeledObjects: 8,
  meaningfulObjects: 3,
  decorativeObjects: 5,
  unclassifiedObjects: 0,
  missingTextObjects: 0,
  opaqueNativeObjects: 0,
});
assert.deepEqual(decorativeAudit.issues, []);
assert.deepEqual(decorativeAudit.manualChecks.map((check) => check.type), ["readingOrder"]);
assert.doesNotMatch(decorativeAudit.ndjson, /"conformanceClaimed"/u, "accessibility audit NDJSON must remain record-only rather than inventing a conformance result");

const decorativeSource = await PresentationFile.exportPptx(decorativePresentation);
const decorativeSourceZip = await JSZip.loadAsync(decorativeSource.bytes);
const decorativeSourceXml = await decorativeSourceZip.file("ppt/slides/slide1.xml").async("text");
assert.equal((decorativeSourceXml.match(/<adec:decorative\b/g) || []).length, 8);
assert.equal((decorativeSourceXml.match(/<adec:decorative\b[^>]*\bval="(?:1|true)"/g) || []).length, 5);
assert.equal((decorativeSourceXml.match(/<adec:decorative\b[^>]*\bval="(?:0|false)"/g) || []).length, 3);
assert.equal((decorativeSourceXml.match(/uri="\{C183D7F6-B498-43B3-948B-1728B52AA6E4\}"/g) || []).length, 8);

const decorativeImported = await PresentationFile.importPptx(decorativeSource);
const decorativeImportedSlide = decorativeImported.slides.getItem(0);
const importedClassifiedShape = itemByName(decorativeImportedSlide.shapes.items, "classified-shape");
const importedDecorativeImage = itemByName(decorativeImportedSlide.images.items, "decorative-image");
const importedDecorativeTable = itemByName(decorativeImportedSlide.tables.items, "decorative-table");
const importedDecorativeChart = itemByName(decorativeImportedSlide.charts.items, "decorative-chart");
const importedDecorativeGroup = itemByName(decorativeImportedSlide.groups.items, "decorative-group");
const importedDecorativeConnector = itemByName(importedDecorativeGroup.connectors.items, "decorative-connector");
assert.deepEqual(importedClassifiedShape.accessibility, { title: "Meaningful status", decorative: false });
for (const object of [importedDecorativeImage, importedDecorativeTable, importedDecorativeChart, importedDecorativeGroup, importedDecorativeConnector]) {
  assert.deepEqual(object.accessibility, { decorative: true });
  assert.deepEqual(object.accessibilityCapability, { sourceBound: true, editable: true, addable: true });
}
assert.equal(importedDecorativeImage.alt, "");
assert.match(decorativeImported.inspect({ kind: "shape,image,table,chart,group,connector", maxChars: 20_000 }).ndjson, /"decorative":true/);
assert.deepEqual(decorativeImported.auditAccessibility().summary, decorativeAudit.summary);
const decorativeNoOp = await PresentationFile.exportPptx(decorativeImported);
assert.deepEqual(decorativeNoOp.bytes, decorativeSource.bytes, "unchanged decorative metadata must return the exact source package");

const decorativeSourceSvg = await decorativeImportedSlide.export({ format: "svg" });
importedClassifiedShape.setAccessibilityMetadata({ title: null, decorative: true });
assert.throws(() => importedDecorativeImage.setAccessibilityMetadata({ title: "Contradictory image" }), /cannot combine decorative: true/i);
assert.deepEqual(importedDecorativeImage.accessibility, { decorative: true }, "a rejected decorative edit must be transactional");
importedDecorativeImage.setAccessibilityMetadata({ decorative: false, title: "Evidence image" });
importedDecorativeTable.setAccessibilityMetadata({ decorative: null, title: "Visual layout key" });
importedDecorativeChart.setAccessibilityMetadata({ decorative: false });
importedDecorativeGroup.setAccessibilityMetadata({ decorative: false, title: "Decorative cluster" });
const decorativeEdited = await PresentationFile.exportPptx(decorativeImported);
const decorativeRoundTrip = await PresentationFile.importPptx(decorativeEdited);
const decorativeRoundTripSlide = decorativeRoundTrip.slides.getItem(0);
assert.deepEqual(itemByName(decorativeRoundTripSlide.shapes.items, "classified-shape").accessibility, { decorative: true });
assert.deepEqual(itemByName(decorativeRoundTripSlide.images.items, "decorative-image").accessibility, { title: "Evidence image", decorative: false });
assert.deepEqual(itemByName(decorativeRoundTripSlide.tables.items, "decorative-table").accessibility, { title: "Visual layout key" });
assert.deepEqual(itemByName(decorativeRoundTripSlide.charts.items, "decorative-chart").accessibility, { decorative: false });
const decorativeRoundTripGroup = itemByName(decorativeRoundTripSlide.groups.items, "decorative-group");
assert.deepEqual(decorativeRoundTripGroup.accessibility, { title: "Decorative cluster", decorative: false });
assert.deepEqual(itemByName(decorativeRoundTripGroup.connectors.items, "decorative-connector").accessibility, { decorative: true });
const decorativeOutputSvg = await decorativeRoundTripSlide.export({ format: "svg" });
assert.deepEqual(decorativeOutputSvg.bytes, decorativeSourceSvg.bytes, "decorative classification must not change model rendering");

const incompleteAccessibilityDeck = Presentation.create();
const incompleteAccessibilitySlide = incompleteAccessibilityDeck.slides.add({ name: "Accessibility audit boundaries" });
incompleteAccessibilityDeck.slides.add({ name: "Empty slide still counted" });
incompleteAccessibilitySlide.shapes.add({
  name: "unclassified-shape",
  position: { left: 20, top: 20, width: 120, height: 60 },
  text: "Needs classification",
});
incompleteAccessibilitySlide.images.add({
  name: "meaningful-without-text",
  position: { left: 180, top: 20, width: 80, height: 60 },
  dataUrl: PNG,
  accessibility: { decorative: false },
});
incompleteAccessibilitySlide.nativeObjects.add({
  name: "opaque-diagram",
  nativeKind: "diagram",
  rawXml: "<p:graphicFrame xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>",
  position: { left: 300, top: 20, width: 120, height: 60 },
});
const incompleteAccessibilityAudit = incompleteAccessibilityDeck.auditAccessibility({ maxChars: 4_000 });
assert.equal(incompleteAccessibilityAudit.machineCheckPassed, false);
assert.equal(incompleteAccessibilityAudit.conformanceClaimed, false);
assert.equal(incompleteAccessibilityAudit.manualReviewRequired, true);
assert.deepEqual(incompleteAccessibilityAudit.issues.map((issue) => issue.type), ["unclassifiedObject", "meaningfulObjectTextMissing"]);
assert.deepEqual(incompleteAccessibilityAudit.manualChecks.map((check) => check.type), ["opaqueObjectAccessibility", "readingOrder"]);
assert.deepEqual(incompleteAccessibilityAudit.summary, {
  slides: 2,
  modeledObjects: 2,
  meaningfulObjects: 1,
  decorativeObjects: 0,
  unclassifiedObjects: 1,
  missingTextObjects: 1,
  opaqueNativeObjects: 1,
});
assert.match(incompleteAccessibilityAudit.ndjson, /shape-tree order would also change visual z-order/u);
assert.throws(() => incompleteAccessibilityDeck.auditAccessibility([]), /options must be an object/i);

// The canonical file facade always crosses the OfficeKit C# WASM layer.
const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const coreSlide = deck.slides.add({ name: "Core objects", background: { fill: "#F1F5F9", mode: "solid" } });
coreSlide.addNotes("Lead with the customer outcome.\nThen explain the operating model.");
coreSlide.shapes.add({
  name: "core-title",
  geometry: "textbox",
  position: { left: 50, top: 28, width: 1180, height: 70 },
  fill: "transparent",
  line: { fill: "transparent", width: 0 },
  text: "Presentation 0.2 core",
  textStyle: { fontFamily: "Arial", fontSize: 38, bold: true, color: "#0F172A" },
});
const rounded = coreSlide.shapes.add({
  name: "rounded-card",
  geometry: "roundRect",
  position: { left: 60, top: 140, width: 260, height: 100 },
  fill: "#DBEAFE",
  line: { fill: "#2563EB", width: 2 },
  shadow: { color: "#000000", blurRadius: 8, distance: 4, direction: 45, opacity: 0.25 },
  text: "Before edit",
  textStyle: { fontFamily: "Arial", fontSize: 25, bold: true, color: "#1E3A8A" },
});
const target = coreSlide.shapes.add({
  name: "target-textbox",
  geometry: "textbox",
  position: { left: 400, top: 140, width: 260, height: 100 },
  fill: "transparent",
  line: { fill: "transparent", width: 0 },
  text: "Target",
  textStyle: { fontFamily: "Arial", fontSize: 25, bold: true, color: "#14532D" },
});
coreSlide.shapes.add({
  name: "rich-copy",
  geometry: "textbox",
  position: { left: 840, top: 270, width: 380, height: 300 },
  fill: "transparent",
  line: { fill: "transparent", width: 0 },
  text: [
    {
      runs: [
        { text: "Structured ", style: { fontSize: 24, bold: true, color: "#0F172A" } },
        { text: "text", style: { fontSize: 24, italic: true, color: "#2563EB" }, link: { uri: "https://www.ecma-international.org/publications-and-standards/standards/ecma-376/" } },
      ],
    },
    {
      bulletCharacter: "•",
      marginLeft: 28,
      indent: -14,
      runs: [{ text: "Character list", style: { fontSize: 19, color: "#334155" } }],
    },
    {
      autoNumber: { type: "arabicPeriod", startAt: 1 },
      marginLeft: 28,
      indent: -14,
      runs: [{ text: "Numbered list", style: { fontSize: 19, color: "#334155" } }],
    },
  ],
});
coreSlide.tables.add({
  name: "fixed-table",
  position: { left: 60, top: 300, width: 360, height: 190 },
  values: [["Layer", "State"], ["Office", "Before"], ["QA", "Ready"]],
  styleOptions: { headerRow: true, bandedRows: true },
});
coreSlide.images.add({
  name: "png-image",
  alt: "PNG evidence",
  position: { left: 470, top: 300, width: 140, height: 140 },
  fit: "stretch",
  dataUrl: PNG,
});
coreSlide.images.add({
  name: "jpeg-image",
  alt: "JPEG evidence",
  position: { left: 650, top: 300, width: 140, height: 140 },
  fit: "stretch",
  dataUrl: JPEG,
});
const coverImage = coreSlide.images.add({
  name: "cover-image",
  alt: "Wide image cropped to a square",
  position: { left: 650, top: 480, width: 140, height: 140 },
  fit: "cover",
  dataUrl: WIDE_SVG,
});
assert.match(coverImage.toSvg(), /viewBox="100 0 200 200"/);
assert.throws(() => { coverImage.crop = { left: 0.8, right: 0.3 }; }, /opposing sums/);
coreSlide.connectors.add({
  name: "straight-connector",
  connectorType: "straight",
  from: rounded,
  to: target,
  start: { x: 320, y: 180 },
  end: { x: 400, y: 180 },
  line: { fill: "#334155", width: 2, endArrow: "triangle" },
});
coreSlide.connectors.add({
  name: "elbow-polyline-connector",
  connectorType: "elbow",
  from: rounded,
  to: target,
  start: { x: 320, y: 210 },
  end: { x: 400, y: 225 },
  line: { fill: "#7C3AED", width: 2, startArrow: "triangle", endArrow: "triangle" },
});

const chartSlide = deck.slides.add({ name: "Literal charts" });
chartSlide.shapes.add({
  name: "chart-title",
  geometry: "textbox",
  position: { left: 50, top: 28, width: 1180, height: 70 },
  fill: "transparent",
  line: { fill: "transparent", width: 0 },
  text: "Source-free bar, line, and pie",
  textStyle: { fontFamily: "Arial", fontSize: 38, bold: true, color: "#0F172A" },
});
chartSlide.charts.add("bar", {
  name: "bar-chart",
  title: "Readiness",
  position: { left: 30, top: 130, width: 380, height: 320 },
  categories: ["Create", "Inspect", "Render"],
  series: [{ name: "Score", values: [78, 92, 85], color: "#2563EB" }],
  legend: false,
  axes: { category: { title: "Gate" }, value: { title: "Score", min: 0, max: 100, majorUnit: 20 } },
  dataLabels: { showValue: true, position: "outsideEnd" },
});
chartSlide.charts.add("line", {
  name: "line-chart",
  title: "Trend",
  position: { left: 450, top: 130, width: 380, height: 320 },
  categories: ["W1", "W2", "W3"],
  series: [{
    name: "Passes",
    values: [6, 9, 12],
    color: "#16A34A",
    line: { fill: "#16A34A", width: 2, style: "dash" },
    marker: { symbol: "circle", size: 7, fill: "#16A34A" },
    trendlines: [
      {
        type: "linear",
        name: "Pass projection",
        forward: 0.5,
        backward: 0.5,
        intercept: 0,
        displayEquation: true,
        displayRSquared: true,
        line: { fill: "#7C3AED", width: 1.5, style: "dash" },
      },
      { type: "movingAverage", name: "Pass moving average", period: 2 },
      { type: "polynomial", name: "Pass curve", order: 2 },
    ],
    errorBars: {
      type: "standardDeviation",
      value: 1.5,
      endStyle: "noCap",
      line: { fill: "#DC2626", width: 1.25, style: "dot" },
    },
  }],
  legend: false,
});
chartSlide.charts.add("pie", {
  name: "pie-chart",
  title: "Coverage",
  position: { left: 870, top: 130, width: 380, height: 320 },
  categories: ["Modeled", "Opaque"],
  series: [{ name: "Share", values: [80, 20], color: "#7C3AED" }],
  legend: true,
  dataLabels: { showCategoryName: true, showValue: true },
});

const comboSlide = deck.slides.add({ name: "Literal combo chart" });
comboSlide.charts.add("combo", {
  name: "revenue-margin-combo",
  title: "Revenue and margin",
  position: { left: 90, top: 120, width: 1080, height: 480 },
  categories: ["Q1", "Q2", "Q3"],
  series: [
    { name: "Revenue", chartType: "bar", values: [42, 48, 57], color: "#2563EB" },
    {
      name: "Margin",
      chartType: "line",
      values: [12, 15, 18],
      color: "#16A34A",
      line: { fill: "#16A34A", width: 2 },
      marker: { symbol: "circle", size: 7, fill: "#16A34A" },
      trendlines: [{ type: "exp", name: "Margin projection", forward: 0.5, line: { fill: "#F97316", width: 1.5, style: "dot" } }],
      errorBars: { type: "minus", valueType: "custom", minusValues: [1, 2, 1], line: { fill: "#EA580C", width: 1 } },
    },
  ],
  legend: true,
  axes: { category: { title: "Quarter" }, value: { title: "Percent" } },
  dataLabels: { showValue: true, position: "top" },
});
assert.match(chartSlide.toSvg(), /data-trendline-type="linear"/);
assert.match(chartSlide.toSvg(), /data-trendline-type="movingAvg"/);
assert.match(chartSlide.toSvg(), /data-trendline-type="poly"/);
assert.match(comboSlide.toSvg(), /data-trendline-type="exp"/);
assert.match(chartSlide.toSvg(), /data-error-bars-series="0"/);
assert.match(comboSlide.toSvg(), /data-error-bars-series="1"/);
const secondaryAxisCombo = Presentation.create({ slideSize: { width: 640, height: 360 } });
const secondaryAxisSlide = secondaryAxisCombo.slides.add({ name: "Secondary-axis combo" });
secondaryAxisSlide.charts.add("combo", {
  name: "secondary-axis-combo",
  title: "Revenue and gross margin",
  position: { left: 48, top: 60, width: 540, height: 250 },
  categories: ["Q1", "Q2"],
  series: [
    { name: "Revenue", chartType: "bar", values: [42, 48], color: "#2563EB" },
    { name: "Gross margin", chartType: "line", axisGroup: "secondary", values: [45, 50], line: { fill: "#16A34A", width: 2 }, marker: { symbol: "circle", size: 6, fill: "#16A34A" } },
  ],
  axes: {
    category: { title: "Quarter" },
    value: { title: "Revenue ($M)" },
    secondary: { category: { title: "Quarter" }, value: { title: "Gross margin (%)", min: 0, max: 100, majorUnit: 10 } },
  },
  legend: true,
});
const secondaryAxisExport = await PresentationFile.exportPptx(secondaryAxisCombo);
const secondaryAxisZip = await JSZip.loadAsync(new Uint8Array(await secondaryAxisExport.arrayBuffer()));
const secondaryAxisChartXml = await Promise.all(Object.keys(secondaryAxisZip.files)
  .filter((name) => /\/charts\/chart\d+\.xml$/.test(name))
  .map((name) => secondaryAxisZip.file(name).async("text")))
  .then((items) => items.find((xml) => xml.includes("Revenue and gross margin")));
assert.ok(secondaryAxisChartXml);
assert.match(secondaryAxisChartXml, /<c:barChart>[\s\S]*?<c:axId val="1"\s*\/><c:axId val="2"\s*\/><\/c:barChart>/);
assert.match(secondaryAxisChartXml, /<c:lineChart>[\s\S]*?<c:axId val="3"\s*\/><c:axId val="4"\s*\/><\/c:lineChart>/);
assert.match(secondaryAxisChartXml, /<c:catAx><c:axId val="3"\s*\/>[\s\S]*?<c:axPos val="t"\s*\/>/);
assert.match(secondaryAxisChartXml, /<c:valAx><c:axId val="4"\s*\/>[\s\S]*?<c:axPos val="r"\s*\/>/);
const importedSecondaryAxis = await PresentationFile.importPptx(secondaryAxisExport);
const importedSecondaryAxisChart = itemByName(importedSecondaryAxis.slides.getItem(0).charts.items, "secondary-axis-combo");
assert.equal(importedSecondaryAxisChart.chartType, "combo");
assert.deepEqual(importedSecondaryAxisChart.series.map((series) => [series.chartType, series.axisGroup || "primary"]), [["bar", "primary"], ["line", "secondary"]]);
assert.equal(importedSecondaryAxisChart.axes.secondary.category.title, "Quarter");
assert.equal(importedSecondaryAxisChart.axes.secondary.value.title, "Gross margin (%)");
assert.equal(importedSecondaryAxisChart.axes.secondary.value.max, 100);
importedSecondaryAxisChart.series[1].values = [47, 53];
importedSecondaryAxisChart.axes.secondary.value.max = 80;
const editedSecondaryAxis = await PresentationFile.exportPptx(importedSecondaryAxis);
const roundTripSecondaryAxis = await PresentationFile.importPptx(editedSecondaryAxis);
const roundTripSecondaryAxisChart = itemByName(roundTripSecondaryAxis.slides.getItem(0).charts.items, "secondary-axis-combo");
assert.deepEqual(roundTripSecondaryAxisChart.series[1].values, [47, 53]);
assert.equal(roundTripSecondaryAxisChart.series[1].axisGroup, "secondary");
assert.equal(roundTripSecondaryAxisChart.axes.secondary.value.max, 80);

const chartFamilyDeck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const chartFamilySlide = chartFamilyDeck.slides.add({ name: "Native chart families" });
chartFamilySlide.charts.add("area", {
  name: "area-family",
  title: "Regional trajectory",
  position: { left: 40, top: 35, width: 570, height: 300 },
  categories: ["Q1", "Q2", "Q3"],
  series: [{ name: "Revenue", values: [42, 53, 68], fill: "#0EA5E9", line: { fill: "#0369A1", width: 1.5 } }],
  xAxis: { title: "Quarter" },
  yAxis: { title: "Revenue", min: 0, max: 80, majorUnit: 20 },
  legend: false,
});
chartFamilySlide.charts.add("doughnut", {
  name: "doughnut-family",
  title: "Regional mix",
  position: { left: 660, top: 35, width: 570, height: 300 },
  categories: ["North", "Central", "South"],
  series: [{ name: "Share", values: [52, 31, 17] }],
  dataLabels: { showCategoryName: true, showPercent: true, position: "outsideEnd" },
  legend: true,
});
chartFamilySlide.charts.add("scatter", {
  name: "scatter-family",
  title: "Reach relationship",
  position: { left: 40, top: 370, width: 570, height: 300 },
  series: [{ name: "Portfolio", xValues: [10, 20, 34], values: [35, 68, 84], marker: { symbol: "diamond", size: 8, fill: "#8B5CF6", line: { fill: "#6D28D9", width: 1 } } }],
  xAxis: { title: "Reach", min: 0, max: 40, majorUnit: 10 },
  yAxis: { title: "Return", min: 0, max: 100, majorUnit: 20 },
  legend: false,
});
chartFamilySlide.charts.add("bubble", {
  name: "bubble-family",
  title: "Opportunity map",
  position: { left: 660, top: 370, width: 570, height: 300 },
  series: [{ name: "Opportunity", xValues: [10, 20, 34], values: [35, 68, 84], bubbleSizes: [4, 9, 16], fill: "#F97316", line: { fill: "#C2410C", width: 1 } }],
  xAxis: { title: "Reach", min: 0, max: 40, majorUnit: 10 },
  yAxis: { title: "Return", min: 0, max: 100, majorUnit: 20 },
  legend: false,
});
assert.equal(chartFamilyDeck.verify().ok, true);
const chartFamilySvg = chartFamilySlide.toSvg();
assert.match(chartFamilySvg, /Regional trajectory/);
assert.match(chartFamilySvg, /52%/);
assert.match(chartFamilySvg, /<circle[^>]+fill-opacity="0\.72"/);
assert.match(chartFamilySvg, /<path[^>]+fill-opacity="0\.45"/);
const chartFamilyExport = await PresentationFile.exportPptx(chartFamilyDeck);
const chartFamilyZip = await JSZip.loadAsync(new Uint8Array(await chartFamilyExport.arrayBuffer()));
const chartFamilyXml = await Promise.all(Object.keys(chartFamilyZip.files)
  .filter((name) => /\/charts\/chart\d+\.xml$/.test(name))
  .map((name) => chartFamilyZip.file(name).async("text")));
assert.equal(chartFamilyXml.filter((xml) => /<c:areaChart>/.test(xml)).length, 1);
assert.equal(chartFamilyXml.filter((xml) => /<c:doughnutChart>/.test(xml)).length, 1);
assert.equal(chartFamilyXml.filter((xml) => /<c:scatterChart>/.test(xml)).length, 1);
assert.equal(chartFamilyXml.filter((xml) => /<c:bubbleChart>/.test(xml)).length, 1);
assert.match(chartFamilyXml.find((xml) => /<c:doughnutChart>/.test(xml)), /<c:showPercent val="1"\s*\/>/);
assert.match(chartFamilyXml.find((xml) => /<c:scatterChart>/.test(xml)), /<c:xVal>[\s\S]*<c:yVal>/);
assert.match(chartFamilyXml.find((xml) => /<c:bubbleChart>/.test(xml)), /<c:xVal>[\s\S]*<c:yVal>[\s\S]*<c:bubbleSize>/);
const importedChartFamilyDeck = await PresentationFile.importPptx(chartFamilyExport);
const importedFamilies = importedChartFamilyDeck.slides.getItem(0).charts.items;
assert.deepEqual(importedFamilies.map((chart) => chart.chartType), ["area", "doughnut", "scatter", "bubble"]);
assert.equal(importedFamilies[1].dataLabels.showPercent, true);
assert.deepEqual(importedFamilies[2].series[0].xValues, [10, 20, 34]);
assert.deepEqual(importedFamilies[3].series[0].bubbleSizes, [4, 9, 16]);
importedFamilies[0].series[0].values[1] = 57;
importedFamilies[1].dataLabels.showPercent = false;
importedFamilies[2].series[0].xValues[1] = 22;
importedFamilies[3].series[0].bubbleSizes[1] = 12;
const editedChartFamilyExport = await PresentationFile.exportPptx(importedChartFamilyDeck);
const roundTripChartFamilies = (await PresentationFile.importPptx(editedChartFamilyExport)).slides.getItem(0).charts.items;
assert.equal(roundTripChartFamilies[0].series[0].values[1], 57);
assert.equal(roundTripChartFamilies[1].dataLabels.showPercent, false);
assert.equal(roundTripChartFamilies[2].series[0].xValues[1], 22);
assert.equal(roundTripChartFamilies[3].series[0].bubbleSizes[1], 12);
assert.throws(() => chartFamilySlide.charts.add("scatter", { categories: ["A"], series: [{ name: "Invalid", xValues: [1], values: [2] }] }), /per-series xValues/i);
assert.throws(() => chartFamilySlide.charts.add("bubble", { series: [{ name: "Invalid", xValues: [1], values: [2], bubbleSizes: [0] }] }), /positive bubbleSize/i);
assert.throws(() => chartFamilySlide.charts.add("doughnut", { categories: ["A"], series: [{ name: "Invalid", values: [1] }], xAxis: { title: "Invalid" } }), /cannot carry axes/i);
assert.throws(() => chartFamilySlide.charts.add("area", { categories: ["A"], series: [{ name: "Invalid marker", values: [1], marker: { symbol: "circle" } }] }), /area series 1 cannot carry a marker/i);
assert.throws(() => chartFamilySlide.charts.add("area", { categories: ["A", "B"], series: [{ name: "Invalid trendline", values: [1, 2], trendlines: [{ type: "linear" }] }] }), /trendlines are supported only for bar and line series/i);
assert.throws(() => chartFamilySlide.charts.add("area", { categories: ["A", "B"], series: [{ name: "Invalid error bars", values: [1, 2], errorBars: { type: "percentage", value: 5 } }] }), /errorBars are supported only for bar and line series/i);
assert.throws(() => chartFamilySlide.charts.add("line", { categories: ["A", "B"], series: [{ name: "Invalid custom error bars", values: [1, 2], errorBars: { valueType: "custom", plusValues: [1, 2] } }] }), /minus requires literal values or a formula/i);
assert.throws(() => chartFamilySlide.charts.add("line", { categories: ["A", "B"], series: [{ name: "Unknown error-bar field", values: [1, 2], errorBars: { type: "percentage", value: 5, confidence: 0.95 } }] }), /unsupported fields: confidence/i);
assert.throws(() => chartFamilySlide.charts.add("line", { categories: ["A", "B"], series: [{ name: "Short average", values: [1, 2], trendlines: [{ type: "movingAvg", period: 2 }] }] }), /require at least three series values/i);
assert.throws(() => chartFamilySlide.charts.add("line", { categories: ["A", "B", "C"], series: [{ name: "Fractional forecast", values: [1, 2, 3], trendlines: [{ type: "linear", forward: 0.25 }] }] }), /must use 0\.5 increments/i);
const formulaErrorBarDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
formulaErrorBarDeck.slides.add().charts.add("line", {
  categories: ["A", "B"],
  externalData: new Uint8Array([0x50, 0x4b, 0x03, 0x04]),
  series: [{
    name: "Formula-backed uncertainty",
    values: [1, 2],
    errorBars: {
      valueType: "custom",
      plusFormula: "Sheet1!$A$1:$A$2",
      minusValues: [0.5, 0.5],
    },
  }],
});
await assert.rejects(
  PresentationFile.exportPptx(formulaErrorBarDeck),
  /series\[0\]\.errorBars\.plusFormula requires an embedded workbook path/i,
);
const horizontalErrorBarDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
const horizontalErrorBarChart = horizontalErrorBarDeck.slides.add().charts.add("bar", {
  categories: ["A", "B"],
  barOptions: { direction: "horizontal" },
  series: [{ name: "Horizontal uncertainty", values: [40, 60], errorBars: { direction: "x", valueType: "fixedVal", value: 10 } }],
});
const horizontalErrorBarMark = /<line data-error-bars-series="0" data-error-bars-index="0" x1="([^"]+)" y1="([^"]+)" x2="([^"]+)" y2="([^"]+)"/.exec(horizontalErrorBarChart.toSvg());
assert.ok(horizontalErrorBarMark);
assert.notEqual(horizontalErrorBarMark[1], horizontalErrorBarMark[3]);
assert.equal(horizontalErrorBarMark[2], horizontalErrorBarMark[4]);
const scatterLineDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
scatterLineDeck.slides.add().charts.add("scatter", { series: [{ name: "Invalid line", xValues: [1, 2], values: [2, 3], line: { fill: "#000000", width: 1 } }] });
await assert.rejects(PresentationFile.exportPptx(scatterLineDeck), /marker-scatter.*cannot carry a series line/i);

const singleAxisChartDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
singleAxisChartDeck.slides.add().charts.add("bar", {
  categories: ["A", "B"],
  series: [{ name: "Values", values: [1, 2] }],
  yAxis: { title: "Configured value axis" },
});
const singleAxisChartRoundTrip = await PresentationFile.importPptx(await PresentationFile.exportPptx(singleAxisChartDeck));
assert.equal(singleAxisChartRoundTrip.slides.getItem(0).charts.items[0].axes.category.title, "");
assert.equal(singleAxisChartRoundTrip.slides.getItem(0).charts.items[0].axes.value.title, "Configured value axis");

const mixedAxisCombo = Presentation.create({ slideSize: { width: 640, height: 360 } });
mixedAxisCombo.slides.add({ name: "Rejected mixed combo" }).charts.add("combo", {
  name: "mixed-axis-combo",
  categories: ["Q1", "Q2"],
  series: [
    { name: "Revenue", chartType: "bar", values: [42, 48] },
    { name: "Primary line", chartType: "line", values: [12, 15] },
    { name: "Secondary line", chartType: "line", axisGroup: "secondary", values: [45, 50] },
  ],
});
await assert.rejects(PresentationFile.exportPptx(mixedAxisCombo), /cannot mix primary and secondary line plots/i);

assert.equal(deck.verify().ok, true);
assert.equal(deck.validateLayout().ok, true);
assert.equal(deck.resolve(rounded.id), rounded);
assert.equal(deck.resolve(rounded.id + "/text").text, "Before edit");
const deckInspect = deck.inspect({ kind: "deck,slide,textbox,shape,table,chart,image,connector,textRange,notes", maxChars: 24_000 }).ndjson;
assert.match(deckInspect, /Lead with the customer outcome/);
assert.match(deckInspect, /"background":\{"fill":"#F1F5F9","mode":"solid"\}/);
assert.equal(deck.resolve(coreSlide.speakerNotes.id), coreSlide.speakerNotes);
coreSlide.speakerNotes.textFrame.setText("Lead with the customer outcome.\nThen explain the operating model.");
assert.equal(coreSlide.speakerNotes.append("").text, "Lead with the customer outcome.\nThen explain the operating model.");

const firstExport = await PresentationFile.exportPptx(deck);
assert.equal(firstExport.metadata.codec, "office-kit");
assert.equal((await PresentationFile.inspectPptx(firstExport)).ok, true);
const changedPresentationTrendlineTopology = await PresentationFile.importPptx(firstExport);
itemByName(changedPresentationTrendlineTopology.slides.getItem(1).charts.items, "line-chart").series[0].trendlines.pop();
await assert.rejects(
  () => PresentationFile.exportPptx(changedPresentationTrendlineTopology),
  (error) => error?.code === "presentation_chart_topology_changed" && /cannot change its imported trendline count/i.test(error.message),
);

const changedPresentationErrorBarTopology = await PresentationFile.importPptx(firstExport);
itemByName(changedPresentationErrorBarTopology.slides.getItem(1).charts.items, "line-chart").series[0].errorBars = undefined;
await assert.rejects(
  () => PresentationFile.exportPptx(changedPresentationErrorBarTopology),
  (error) => error?.code === "presentation_chart_topology_changed" && /cannot add or remove imported error bars/i.test(error.message),
);

// Speaker notes use the same paragraph/run model as visible slide text, but
// retain a deliberately narrower relationship-free contract. This proves the
// public facade can author, reimport, and edit a multi-run talk track without
// flattening it through the legacy `.text` convenience field.
const richNotesDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
const richNotesSlide = richNotesDeck.slides.add({
  name: "Rich speaker notes",
  notes: [
    {
      bulletCharacter: "•",
      runs: [
        { text: "Open with ", style: { bold: true, fontSize: 18, fontFamily: "Aptos", color: "#0F172A" } },
        { text: "the customer outcome.", style: { italic: true, fontSize: 18 } },
      ],
    },
    { autoNumber: { type: "arabicPeriod", startAt: 2 }, runs: [{ text: "Then explain the operating model.", style: { fontSize: 16 } }] },
  ],
});
richNotesSlide.shapes.add({ name: "rich-notes-title", text: "Visible slide", position: { left: 48, top: 48, width: 300, height: 72 } });
const richNotesPptx = await PresentationFile.exportPptx(richNotesDeck);
const richNotesZip = await JSZip.loadAsync(richNotesPptx.bytes);
const richNotesXml = await richNotesZip.file("ppt/notesSlides/notesSlide1.xml").async("text");
assert.match(richNotesXml, /<a:buChar\b[^>]*char="•"/);
assert.match(richNotesXml, /<a:rPr\b[^>]*\bb="1"/);
assert.match(richNotesXml, /<a:rPr\b[^>]*\bi="1"/);
assert.match(richNotesXml, /<a:buAutoNum\b[^>]*type="arabicPeriod"[^>]*startAt="2"/);
const importedRichNotesDeck = await PresentationFile.importPptx(richNotesPptx);
const importedRichNotes = importedRichNotesDeck.slides.getItem(0).speakerNotes;
assert.equal(importedRichNotes.text, "Open with the customer outcome.\nThen explain the operating model.");
assert.equal(importedRichNotes.capability.editable, true);
const importedRichParagraphs = importedRichNotes.textFrame.paragraphs;
assert.equal(importedRichParagraphs.length, 2);
assert.equal(importedRichParagraphs[0].runs.length, 2);
assert.equal(importedRichParagraphs[0].runs[0].style.bold, true);
assert.equal(importedRichParagraphs[0].runs[1].style.italic, true);
assert.deepEqual(importedRichParagraphs[1].autoNumber, { type: "arabicPeriod", startAt: 2 });
const richNotesNoOpPptx = await PresentationFile.exportPptx(importedRichNotesDeck);
const richNotesNoOpZip = await JSZip.loadAsync(richNotesNoOpPptx.bytes);
assert.deepEqual(
  await richNotesNoOpZip.file("ppt/notesSlides/notesSlide1.xml").async("uint8array"),
  await richNotesZip.file("ppt/notesSlides/notesSlide1.xml").async("uint8array"),
  "unchanged imported rich notes must retain their source NotesSlide bytes",
);
importedRichParagraphs[0].runs[1].text = "the operating decision.";
importedRichParagraphs[0].runs[1].style = { ...importedRichParagraphs[0].runs[1].style, bold: true, italic: false };
importedRichNotes.textFrame.paragraphs = importedRichParagraphs;
const richNotesEditedPptx = await PresentationFile.exportPptx(importedRichNotesDeck);
const richNotesEditedDeck = await PresentationFile.importPptx(richNotesEditedPptx);
const editedRichNotes = richNotesEditedDeck.slides.getItem(0).speakerNotes;
assert.equal(editedRichNotes.text, "Open with the operating decision.\nThen explain the operating model.");
assert.equal(editedRichNotes.textFrame.paragraphs[0].runs[1].style.bold, true);
assert.equal(editedRichNotes.textFrame.paragraphs[0].runs[1].style.italic, false);
const richNotesFlattenAttempt = await PresentationFile.importPptx(richNotesPptx);
richNotesFlattenAttempt.slides.getItem(0).speakerNotes.text = "Do not flatten this multi-run talk track.";
await assert.rejects(
  () => PresentationFile.exportPptx(richNotesFlattenAttempt),
  (error) => error?.code === "presentation_text_topology_changed",
);

// Imported deck reordering is intentionally a shallow package operation: it
// preserves every original SlidePart exactly once and changes only the
// p:sldIdLst display order. It is separate from the graph-ownership-based
// imported-slide deletion profile below.
const reorderedImportedDeck = await PresentationFile.importPptx(firstExport);
const originalImportedSlideNames = reorderedImportedDeck.slides.items.map((slide) => slide.name);
const importedFirstSlide = reorderedImportedDeck.slides.getItem(0);
assert.equal(importedFirstSlide.moveTo(2), importedFirstSlide);
assert.deepEqual(reorderedImportedDeck.slides.items.map((slide) => slide.name), [...originalImportedSlideNames.slice(1), originalImportedSlideNames[0]]);
assert.throws(
  () => importedFirstSlide.moveTo(3),
  /destination must be an existing 0-based slide index/i,
);
const reorderedImportedPptx = await PresentationFile.exportPptx(reorderedImportedDeck);
const originalImportedZip = await JSZip.loadAsync(firstExport.bytes);
const reorderedImportedZip = await JSZip.loadAsync(reorderedImportedPptx.bytes);
for (const path of Object.keys(originalImportedZip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))) {
  assert.deepEqual(
    await reorderedImportedZip.file(path).async("uint8array"),
    await originalImportedZip.file(path).async("uint8array"),
    `reordering must preserve ${path} byte-for-byte`,
  );
}
const reorderedImportedRoundTrip = await PresentationFile.importPptx(reorderedImportedPptx);
assert.deepEqual(reorderedImportedRoundTrip.slides.items.map((slide) => slide.name), [...originalImportedSlideNames.slice(1), originalImportedSlideNames[0]]);

// A retained imported SlidePart may change only its native p:cSld/@name. The
// transaction must leave every other decoded package part byte-for-byte intact.
const renamedImportedDeck = await PresentationFile.importPptx(firstExport);
const renamedImportedSlide = renamedImportedDeck.slides.getItem(0);
renamedImportedSlide.name = "Renamed imported overview";
const renamedImportedPptx = await PresentationFile.exportPptx(renamedImportedDeck);
const renamedImportedZip = await JSZip.loadAsync(renamedImportedPptx.bytes);
assert.deepEqual(Object.keys(renamedImportedZip.files).sort(), Object.keys(originalImportedZip.files).sort());
for (const [path, entry] of Object.entries(originalImportedZip.files)) {
  if (entry.dir || path === "ppt/slides/slide1.xml") continue;
  assert.deepEqual(
    await renamedImportedZip.file(path).async("uint8array"),
    await originalImportedZip.file(path).async("uint8array"),
    `renaming an imported slide must preserve ${path} byte-for-byte`,
  );
}
assert.match(await renamedImportedZip.file("ppt/slides/slide1.xml").async("text"), /<p:cSld\b[^>]*\bname="Renamed imported overview"/);
const renamedImportedRoundTrip = await PresentationFile.importPptx(renamedImportedPptx);
assert.equal(renamedImportedRoundTrip.slides.getItem(0).name, "Renamed imported overview");
assert.equal(itemByName(renamedImportedRoundTrip.slides.getItem(0).shapes.items, "rounded-card").text.value, "Before edit");

// An imported canvas resize changes only p:presentation/p:sldSz. It must not
// silently rescale every coordinate in a source-bound deck; agents can make a
// separate, explicit layout decision after selecting the new canvas.
const resizedImportedDeck = await PresentationFile.importPptx(firstExport);
resizedImportedDeck.slideSize = { width: 960, height: 720 };
const resizedImportedPptx = await PresentationFile.exportPptx(resizedImportedDeck);
const resizedImportedZip = await JSZip.loadAsync(resizedImportedPptx.bytes);
assert.deepEqual(Object.keys(resizedImportedZip.files).sort(), Object.keys(originalImportedZip.files).sort());
for (const [path, entry] of Object.entries(originalImportedZip.files)) {
  if (entry.dir || path === "ppt/presentation.xml") continue;
  assert.deepEqual(
    await resizedImportedZip.file(path).async("uint8array"),
    await originalImportedZip.file(path).async("uint8array"),
    `resizing an imported canvas must preserve ${path} byte-for-byte`,
  );
}
const resizedPresentationXml = await resizedImportedZip.file("ppt/presentation.xml").async("text");
assert.match(resizedPresentationXml, /<p:sldSz\b[^>]*\bcx="9144000"[^>]*\bcy="6858000"/);
assert.doesNotMatch(resizedPresentationXml, /<p:sldSz\b[^>]*\btype=/);
const resizedImportedRoundTrip = await PresentationFile.importPptx(resizedImportedPptx);
assert.deepEqual(resizedImportedRoundTrip.slideSize, { width: 960, height: 720 });
assert.equal(itemByName(resizedImportedRoundTrip.slides.getItem(0).shapes.items, "rounded-card").text.value, "Before edit");
const invalidImportedCanvas = await PresentationFile.importPptx(firstExport);
invalidImportedCanvas.slideSize = { width: 0, height: 720 };
await assert.rejects(
  () => PresentationFile.exportPptx(invalidImportedCanvas),
  (error) => error?.code === "invalid_slide_size",
);

const reorderedEditedDeck = await PresentationFile.importPptx(firstExport);
const reorderedEditedSlide = reorderedEditedDeck.slides.getItem(0);
reorderedEditedSlide.moveTo(2);
itemByName(reorderedEditedSlide.shapes.items, "rounded-card").text.set("Edited after reorder");
const reorderedEditedPptx = await PresentationFile.exportPptx(reorderedEditedDeck);
const reorderedEditedRoundTrip = await PresentationFile.importPptx(reorderedEditedPptx);
assert.equal(itemByName(reorderedEditedRoundTrip.slides.getItem(2).shapes.items, "rounded-card").text.value, "Edited after reorder");
const importedTopologyChange = await PresentationFile.importPptx(firstExport);
importedTopologyChange.slides.add({ name: "Not source-bound" });
await assert.rejects(
  () => PresentationFile.exportPptx(importedTopologyChange),
  (error) => error?.code === "presentation_topology_changed",
);

// Imported deletion is a real OPC delete, not hiding a slide from p:sldIdLst.
// The bounded profile accepts an otherwise isolated SlidePart with only its
// layout relation, preserves every retained source part byte-for-byte, and
// refuses to remove the final remaining slide before a package write begins.
const deletionFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const deletionKeep = deletionFixture.slides.add({ name: "Keep" });
deletionKeep.shapes.add({ name: "keep-copy", position: { left: 48, top: 48, width: 300, height: 72 }, text: "Keep" });
const deletionRemove = deletionFixture.slides.add({ name: "Remove" });
deletionRemove.shapes.add({ name: "remove-copy", position: { left: 48, top: 48, width: 300, height: 72 }, text: "Remove" });
assert.deepEqual(deletionRemove.deletionCapability, {
  sourceBound: false,
  known: true,
  supported: true,
  blockedReason: "",
  ownedPartCount: 0,
});
const deletionSourcePptx = await PresentationFile.exportPptx(deletionFixture);
const deletionSourceZip = await JSZip.loadAsync(deletionSourcePptx.bytes);
const deletionImportedDeck = await PresentationFile.importPptx(deletionSourcePptx);
const deletionImportedSlide = deletionImportedDeck.slides.getItem(1);
assert.deepEqual(deletionImportedSlide.deletionCapability, {
  sourceBound: true,
  known: true,
  supported: true,
  blockedReason: "",
  ownedPartCount: 1,
});
assert.match(deletionImportedDeck.inspect({ kind: "slide" }).ndjson, /"deletionCapability":\{"sourceBound":true,"known":true,"supported":true/);
assert.equal(deletionImportedSlide.delete(), undefined);
assert.throws(() => deletionImportedDeck.slides.getItem(0).delete(), /retain at least one slide/i);
const deletionPptx = await PresentationFile.exportPptx(deletionImportedDeck);
const deletionZip = await JSZip.loadAsync(deletionPptx.bytes);
assert.equal(deletionZip.file("ppt/slides/slide2.xml"), null, "the deleted slide part must not remain in the package");
assert.equal(deletionZip.file("ppt/slides/_rels/slide2.xml.rels"), null, "the deleted slide relationship part must not remain in the package");
assert.deepEqual(
  await deletionZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await deletionSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "deleting another imported slide must retain the survivor byte-for-byte",
);
const deletionRoundTrip = await PresentationFile.importPptx(deletionPptx);
assert.deepEqual(deletionRoundTrip.slides.items.map((slide) => slide.name), ["Keep"]);

// Imported duplication is a distinct OPC graph operation, not another
// p:sldId reference to the same SlidePart. The first profile deliberately
// stays small: an unchanged shape-only slide with its layout as the only
// relationship becomes a fresh part; after export/reimport it is an ordinary
// source-bound slide and can use the normal supported edit path.
assert.throws(
  () => deletionFixture.slides.getItem(0).duplicate(),
  /available only for a supported imported PPTX source slide/i,
);
const cloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const cloneOriginal = cloneFixture.slides.add({ name: "Original" });
cloneOriginal.shapes.add({ name: "clone-copy", position: { left: 48, top: 48, width: 300, height: 72 }, text: "Original" });
cloneFixture.slides.add({ name: "Companion" }).shapes.add({ name: "companion-copy", position: { left: 48, top: 48, width: 300, height: 72 }, text: "Companion" });
const cloneSourcePptx = await PresentationFile.exportPptx(cloneFixture);
const cloneSourceZip = await JSZip.loadAsync(cloneSourcePptx.bytes);
const cloneImportedDeck = await PresentationFile.importPptx(cloneSourcePptx);
const importedCloneSource = cloneImportedDeck.slides.getItem(0);
const importedClone = importedCloneSource.duplicate();
assert.equal(importedClone.index, 1);
assert.notEqual(importedClone.id, importedCloneSource.id);
assert.notEqual(importedClone.shapes.items[0].id, importedCloneSource.shapes.items[0].id);
assert.equal(importedClone.shapes.items[0].text.value, "Original");
assert.throws(
  () => importedCloneSource.duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);
const clonePptx = await PresentationFile.exportPptx(cloneImportedDeck);
const cloneZip = await JSZip.loadAsync(clonePptx.bytes);
const cloneSlidePath = (await orderedPptxSlidePaths(cloneZip))[1];
assert.deepEqual(
  await cloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await cloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning an imported slide must leave its origin SlidePart byte-for-byte intact",
);
assert.ok(cloneZip.file(cloneSlidePath), "the clone must be a new SlidePart rather than another reference to slide1");
const cloneRoundTrip = await PresentationFile.importPptx(clonePptx);
assert.deepEqual(cloneRoundTrip.slides.items.map((slide) => slide.name), ["Original", "Original", "Companion"]);
itemByName(cloneRoundTrip.slides.getItem(1).shapes.items, "clone-copy").text.set("Edited after reimport");
const cloneEditedPptx = await PresentationFile.exportPptx(cloneRoundTrip);
const cloneEditedRoundTrip = await PresentationFile.importPptx(cloneEditedPptx);
assert.equal(itemByName(cloneEditedRoundTrip.slides.getItem(1).shapes.items, "clone-copy").text.value, "Edited after reimport");

// Clone eligibility is now an OPC ownership proof rather than a catalog of
// known PresentationML object types. Unknown, exclusively owned descendants
// and their external relationships are copied exactly; a node shared by two
// source slides is rejected before the JavaScript model is mutated.
const cloneBaseSlide1Relationships = await cloneSourceZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const cloneBaseSlide2Relationships = await cloneSourceZip.file("ppt/slides/_rels/slide2.xml.rels").async("text");
const ownedGraphSource = await PresentationFile.patchPptx(cloneSourcePptx, [
  {
    path: "ppt/slides/_rels/slide1.xml.rels",
    xml: cloneBaseSlide1Relationships.replace("</Relationships>", '<Relationship Id="rIdAgentOwnedRoot" Type="urn:office-kit:test/owned-root" Target="../customXml/agent-owned-root.xml"/></Relationships>'),
  },
  { path: "ppt/customXml/agent-owned-root.xml", contentType: "application/vnd.office-kit.test+xml", xml: '<owned xmlns="urn:office-kit:test">agent-owned-root</owned>' },
  {
    path: "ppt/customXml/_rels/agent-owned-root.xml.rels",
    xml: '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdAgentOwnedLeaf" Type="urn:office-kit:test/owned-leaf" Target="agent-owned-leaf.bin"/><Relationship Id="rIdAgentOwnedExternal" Type="urn:office-kit:test/owned-external" Target="https://example.invalid/opaque-resource" TargetMode="External"/></Relationships>',
  },
  { path: "ppt/customXml/agent-owned-leaf.bin", contentType: "application/vnd.office-kit.test", bytes: Uint8Array.from([9, 7, 5, 3, 1]) },
]);
const ownedGraphDeck = await PresentationFile.importPptx(ownedGraphSource);
assert.deepEqual(ownedGraphDeck.slides.getItem(0).cloneCapability, {
  sourceBound: true,
  known: true,
  supported: true,
  blockedReason: "",
  clonedPartCount: 3,
  sharedPartCount: 1,
});
ownedGraphDeck.slides.getItem(0).duplicate();
const ownedGraphOutput = await PresentationFile.exportPptx(ownedGraphDeck);
const ownedGraphOutputZip = await JSZip.loadAsync(ownedGraphOutput.bytes);
const ownedRootPaths = Object.keys(ownedGraphOutputZip.files).filter((partPath) => /^ppt\/customXml\/[^/]*owned-root[^/]*\.xml$/i.test(partPath));
const ownedLeafPaths = Object.keys(ownedGraphOutputZip.files).filter((partPath) => /^ppt\/customXml\/[^/]*owned-leaf[^/]*\.bin$/i.test(partPath));
assert.equal(ownedRootPaths.length, 2);
assert.equal(ownedLeafPaths.length, 2);
assert.deepEqual(await ownedGraphOutputZip.file(ownedRootPaths[0]).async("uint8array"), await ownedGraphOutputZip.file(ownedRootPaths[1]).async("uint8array"));
assert.deepEqual(await ownedGraphOutputZip.file(ownedLeafPaths[0]).async("uint8array"), await ownedGraphOutputZip.file(ownedLeafPaths[1]).async("uint8array"));
for (const rootPath of ownedRootPaths) assert.match(await ownedGraphOutputZip.file(relationshipPartPath(rootPath)).async("text"), /rIdAgentOwnedExternal/);

const sharedGraphSource = await PresentationFile.patchPptx(cloneSourcePptx, [
  {
    path: "ppt/slides/_rels/slide1.xml.rels",
    xml: cloneBaseSlide1Relationships.replace("</Relationships>", '<Relationship Id="rIdAgentShared" Type="urn:office-kit:test/shared" Target="../customXml/agent-shared.xml"/></Relationships>'),
  },
  {
    path: "ppt/slides/_rels/slide2.xml.rels",
    xml: cloneBaseSlide2Relationships.replace("</Relationships>", '<Relationship Id="rIdAgentShared" Type="urn:office-kit:test/shared" Target="../customXml/agent-shared.xml"/></Relationships>'),
  },
  { path: "ppt/customXml/agent-shared.xml", contentType: "application/vnd.office-kit.shared+xml", xml: '<shared xmlns="urn:office-kit:test">agent-shared</shared>' },
]);
const sharedGraphDeck = await PresentationFile.importPptx(sharedGraphSource);
assert.equal(sharedGraphDeck.slides.getItem(0).cloneCapability.supported, false);
assert.match(sharedGraphDeck.slides.getItem(0).cloneCapability.blockedReason, /also referenced/);
assert.throws(() => sharedGraphDeck.slides.getItem(0).duplicate(), (error) => error?.code === "unsupported_presentation_slide_clone");
assert.equal(sharedGraphDeck.slides.items.length, 2);

// Canonical run hyperlinks are a modeled part of the same bounded clone leaf.
// The clone keeps the source XML r:ids, creates equivalent external/internal
// relationships on its fresh SlidePart, and retains action-only links inline.
const hyperlinkCloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const hyperlinkCloneOrigin = hyperlinkCloneFixture.slides.add({ name: "Linked origin" });
const hyperlinkCloneTarget = hyperlinkCloneFixture.slides.add({ name: "Linked target" });
hyperlinkCloneTarget.shapes.add({ name: "target-copy", position: { left: 48, top: 48, width: 300, height: 72 }, text: "Target" });
hyperlinkCloneFixture.slides.add({ name: "Linked appendix" });
hyperlinkCloneOrigin.shapes.add({
  name: "linked-copy",
  geometry: "textbox",
  position: { left: 48, top: 48, width: 520, height: 96 },
  fill: "transparent",
  line: { fill: "transparent", width: 0 },
  text: [{
    runs: [
      { text: "Guide ", link: { uri: "https://example.com/guide?x=1&y=2", tooltip: "Read the guide", targetFrame: "_blank" } },
      { text: "Target ", link: { slideId: hyperlinkCloneTarget.id, tooltip: "Open target" } },
      { text: "Next", link: { action: "nextSlide" } },
    ],
  }],
});
const hyperlinkCloneSourcePptx = await PresentationFile.exportPptx(hyperlinkCloneFixture);
const hyperlinkCloneSourceZip = await JSZip.loadAsync(hyperlinkCloneSourcePptx.bytes);
const hyperlinkCloneImported = await PresentationFile.importPptx(hyperlinkCloneSourcePptx);
const hyperlinkCloneImportedOrigin = hyperlinkCloneImported.slides.getItem(0);
const hyperlinkCloneImportedTarget = hyperlinkCloneImported.slides.getItem(1);
assert.equal(hyperlinkCloneImportedTarget.deletionCapability.sourceBound, true);
assert.equal(hyperlinkCloneImportedTarget.deletionCapability.known, true);
assert.equal(hyperlinkCloneImportedTarget.deletionCapability.supported, false);
assert.match(hyperlinkCloneImportedTarget.deletionCapability.blockedReason, /referenced/i);
const hyperlinkSlideCountBeforeRejectedDelete = hyperlinkCloneImported.slides.count;
assert.throws(
  () => hyperlinkCloneImportedTarget.delete(),
  (error) => error?.code === "unsupported_presentation_slide_delete",
);
assert.equal(hyperlinkCloneImported.slides.count, hyperlinkSlideCountBeforeRejectedDelete, "failed preflight must not mutate slide topology");
const hyperlinkClonePending = hyperlinkCloneImportedOrigin.duplicate();
const pendingRuns = itemByName(hyperlinkClonePending.shapes.items, "linked-copy").text.paragraphs[0].runs;
assert.equal(pendingRuns[0].link.uri, "https://example.com/guide?x=1&y=2");
assert.equal(pendingRuns[1].link.slideId, hyperlinkCloneImported.slides.getItem(2).id);
assert.equal(pendingRuns[2].link.action, "nextSlide");
const hyperlinkClonePptx = await PresentationFile.exportPptx(hyperlinkCloneImported);
const hyperlinkCloneZip = await JSZip.loadAsync(hyperlinkClonePptx.bytes);
const hyperlinkCloneSlidePath = (await orderedPptxSlidePaths(hyperlinkCloneZip))[1];
assert.deepEqual(
  await hyperlinkCloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await hyperlinkCloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning canonical run hyperlinks must retain the origin SlidePart byte-for-byte",
);
assert.deepEqual(
  await hyperlinkCloneZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  await hyperlinkCloneSourceZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  "cloning canonical run hyperlinks must retain the origin relationship part byte-for-byte",
);
const modeledRunLinkRelationships = (relationships) => [...relationships.matchAll(/<Relationship\b[^>]*>/g)]
  .map(([tag]) => ({
    id: /\bId="([^"]+)"/.exec(tag)?.[1],
    type: /\bType="([^"]+)"/.exec(tag)?.[1],
    target: /\bTarget="([^"]+)"/.exec(tag)?.[1],
    targetMode: /\bTargetMode="([^"]+)"/.exec(tag)?.[1],
  }))
  .filter((relationship) => /\/(?:hyperlink|slide)$/.test(relationship.type || ""))
  .sort((left, right) => left.id.localeCompare(right.id));
assert.deepEqual(
  modeledRunLinkRelationships(await hyperlinkCloneZip.file(relationshipPartPath(hyperlinkCloneSlidePath)).async("text")),
  modeledRunLinkRelationships(await hyperlinkCloneZip.file("ppt/slides/_rels/slide1.xml.rels").async("text")),
  "the clone must own the same canonical external and internal run-link graph",
);
const hyperlinkCloneRoundTrip = await PresentationFile.importPptx(hyperlinkClonePptx);
const hyperlinkCloneShape = itemByName(hyperlinkCloneRoundTrip.slides.getItem(1).shapes.items, "linked-copy");
const roundTripRuns = hyperlinkCloneShape.text.paragraphs[0].runs;
assert.equal(roundTripRuns[0].link.uri, "https://example.com/guide?x=1&y=2");
assert.equal(roundTripRuns[1].link.slideId, hyperlinkCloneRoundTrip.slides.getItem(2).id);
assert.equal(roundTripRuns[2].link.action, "nextSlide");
roundTripRuns[0].link = { uri: "https://example.com/clone-updated" };
hyperlinkCloneShape.text.paragraphs = [{ runs: roundTripRuns }];
const hyperlinkCloneEdited = await PresentationFile.exportPptx(hyperlinkCloneRoundTrip);
const hyperlinkCloneEditedRoundTrip = await PresentationFile.importPptx(hyperlinkCloneEdited);
assert.equal(itemByName(hyperlinkCloneEditedRoundTrip.slides.getItem(0).shapes.items, "linked-copy").text.paragraphs[0].runs[0].link.uri, "https://example.com/guide?x=1&y=2");
assert.equal(itemByName(hyperlinkCloneEditedRoundTrip.slides.getItem(1).shapes.items, "linked-copy").text.paragraphs[0].runs[0].link.uri, "https://example.com/clone-updated");

const immediateHyperlinkCloneEdit = await PresentationFile.importPptx(hyperlinkCloneSourcePptx);
const immediateHyperlinkCloneShape = itemByName(immediateHyperlinkCloneEdit.slides.getItem(0).duplicate().shapes.items, "linked-copy");
const immediateHyperlinkRuns = immediateHyperlinkCloneShape.text.paragraphs[0].runs;
immediateHyperlinkRuns[0].link = { uri: "https://example.com/too-soon" };
immediateHyperlinkCloneShape.text.paragraphs = [{ runs: immediateHyperlinkRuns }];
await assert.rejects(
  () => PresentationFile.exportPptx(immediateHyperlinkCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

const immediateCloneEdit = await PresentationFile.importPptx(cloneSourcePptx);
immediateCloneEdit.slides.getItem(0).duplicate().shapes.items[0].text.set("Too soon");
await assert.rejects(
  () => PresentationFile.exportPptx(immediateCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

const immediateCloneRename = await PresentationFile.importPptx(cloneSourcePptx);
immediateCloneRename.slides.getItem(0).duplicate().name = "Too soon";
await assert.rejects(
  () => PresentationFile.exportPptx(immediateCloneRename),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

const cloneWithoutOrigin = await PresentationFile.importPptx(cloneSourcePptx);
const cloneOrigin = cloneWithoutOrigin.slides.getItem(0);
cloneOrigin.duplicate();
cloneOrigin.delete();
await assert.rejects(
  () => PresentationFile.exportPptx(cloneWithoutOrigin),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

// The next clone leaf includes canonical embedded images. The fresh SlidePart
// owns a fresh relationship, but both slides deliberately point at the same
// immutable media part; no source slide XML or media bytes are rewritten.
const imageCloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const imageCloneOriginal = imageCloneFixture.slides.add({ name: "Image original" });
imageCloneOriginal.shapes.add({ name: "image-clone-copy", position: { left: 48, top: 48, width: 300, height: 72 }, text: "Image original" });
imageCloneOriginal.images.add({
  name: "image-clone-asset",
  alt: "Shared immutable clone asset",
  position: { left: 48, top: 150, width: 120, height: 120 },
  dataUrl: PNG,
  fit: "stretch",
});
const imageCloneSourcePptx = await PresentationFile.exportPptx(imageCloneFixture);
const imageCloneSourceZip = await JSZip.loadAsync(imageCloneSourcePptx.bytes);
const imageCloneSourceMediaPaths = Object.keys(imageCloneSourceZip.files)
  .filter((path) => /^ppt\/media\/[^/]+\.(?:png|jpe?g|gif|svg)$/i.test(path));
assert.equal(imageCloneSourceMediaPaths.length, 1, "the source fixture must contain exactly one embedded image part");
const [imageCloneSourceMediaPath] = imageCloneSourceMediaPaths;
const imageCloneImportedDeck = await PresentationFile.importPptx(imageCloneSourcePptx);
const imageCloneImportedSource = imageCloneImportedDeck.slides.getItem(0);
const imageClone = imageCloneImportedSource.duplicate();
assert.equal(imageClone.images.items.length, 1);
assert.notEqual(imageClone.images.items[0].id, imageCloneImportedSource.images.items[0].id);
assert.equal(imageClone.images.items[0].alt, "Shared immutable clone asset");
assert.equal(imageClone.images.items[0].dataUrl, PNG);
const imageClonePptx = await PresentationFile.exportPptx(imageCloneImportedDeck);
const imageCloneZip = await JSZip.loadAsync(imageClonePptx.bytes);
const imageCloneSlidePath = (await orderedPptxSlidePaths(imageCloneZip))[1];
assert.deepEqual(
  await imageCloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await imageCloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning a slide with an embedded image must retain the origin SlidePart byte-for-byte",
);
assert.deepEqual(
  await imageCloneZip.file(imageCloneSourceMediaPath).async("uint8array"),
  await imageCloneSourceZip.file(imageCloneSourceMediaPath).async("uint8array"),
  "cloning a slide with an embedded image must retain the shared media bytes",
);
assert.ok(imageCloneZip.file(imageCloneSlidePath), "the image clone must own a new SlidePart");
const imageParts = Object.keys(imageCloneZip.files).filter((path) => /^ppt\/media\/[^/]+\.(?:png|jpe?g|gif|svg)$/i.test(path));
assert.deepEqual(imageParts, [imageCloneSourceMediaPath], "the clone must reuse the source ImagePart instead of duplicating media bytes");
const imageRelationshipTargets = (relationships) => [...relationships.matchAll(/<Relationship\b[^>]*>/g)]
  .filter(([tag]) => /\bType="[^\"]*\/image"/.test(tag))
  .map(([tag]) => /\bTarget="([^\"]+)"/.exec(tag)?.[1]);
const sourceImageRelationships = await imageCloneZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const cloneImageRelationships = await imageCloneZip.file(relationshipPartPath(imageCloneSlidePath)).async("text");
assert.deepEqual(
  imageRelationshipTargets(cloneImageRelationships),
  imageRelationshipTargets(sourceImageRelationships),
  "the clone must own equivalent image relationships to the same media targets",
);
const imageCloneRoundTrip = await PresentationFile.importPptx(imageClonePptx);
assert.deepEqual(imageCloneRoundTrip.slides.items.map((slide) => slide.name), ["Image original", "Image original"]);
assert.equal(imageCloneRoundTrip.slides.getItem(1).images.items[0].alt, "Shared immutable clone asset");
imageCloneRoundTrip.slides.getItem(1).images.items[0].alt = "Edited after image clone reimport";
const imageCloneEditedPptx = await PresentationFile.exportPptx(imageCloneRoundTrip);
const imageCloneEditedRoundTrip = await PresentationFile.importPptx(imageCloneEditedPptx);
assert.equal(imageCloneEditedRoundTrip.slides.getItem(1).images.items[0].alt, "Edited after image clone reimport");

const immediateImageCloneEdit = await PresentationFile.importPptx(imageCloneSourcePptx);
immediateImageCloneEdit.slides.getItem(0).duplicate().images.items[0].alt = "Too soon";
await assert.rejects(
  () => PresentationFile.exportPptx(immediateImageCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

// Canonical tables are an accepted GraphicFrame leaf. They are inline in slide
// XML, so duplicating them must create fresh model identity and exactly no
// table-specific OPC relationship. Closed literal-data charts are exercised
// separately below because each clone must own a distinct ChartPart.
const tableCloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const tableCloneOriginal = tableCloneFixture.slides.add({ name: "Table original" });
const tableCloneSourceTable = tableCloneOriginal.tables.add({
  name: "decision-grid",
  position: { left: 48, top: 48, width: 450, height: 210 },
  values: [["Release evidence", "discarded", "discarded"], ["Native QA", "discarded", "Pass"], ["discarded", "discarded", "Release"]],
  rows: 3,
  columns: 3,
  styleOptions: { headerRow: true, bandedRows: true },
});
assert.equal(tableCloneSourceTable.merge({ startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 }), tableCloneSourceTable);
tableCloneSourceTable.merge({ startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 });
assert.deepEqual(tableCloneSourceTable.mergeRanges, [
  { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
  { startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 },
]);
assert.deepEqual(tableCloneSourceTable.values, [["Release evidence", "", ""], ["Native QA", "", "Pass"], ["", "", "Release"]]);
assert.equal(tableCloneSourceTable.getCell(0, 0).columnSpan, 3);
assert.equal(tableCloneSourceTable.getCell(1, 0).rowSpan, 2);
assert.deepEqual(tableCloneSourceTable.getCell(2, 1).mergeOrigin, { row: 1, column: 0 });
assert.equal(tableCloneSourceTable.getCell(2, 1).editable, false);
assert.throws(() => tableCloneSourceTable.cells.set(2, 1, "hidden"), /covered by merge origin 1,0.*read-only/i);
assert.throws(() => tableCloneSourceTable.merge({ startRow: 0, endRow: 1, startColumn: 2, endColumn: 2 }), /overlap at cell 0,2/i);
assert.throws(() => tableCloneSourceTable.merge({ startRow: 3, endRow: 3, startColumn: 0, endColumn: 1 }), /outside the 3x3 grid/i);
assert.throws(() => tableCloneSourceTable.merge({ startRow: 2, endRow: 2, startColumn: 2, endColumn: 2 }), /at least two cells/i);
assert.match(tableCloneSourceTable.toSvg(), /width="450" height="70"/);
assert.match(tableCloneSourceTable.toSvg(), /width="300" height="140"/);
assert.match(tableCloneFixture.inspect({ kind: "table" }).ndjson, /"mergeRanges"/);
assert.deepEqual(tableCloneSourceTable.layoutJson().mergeRanges, tableCloneSourceTable.mergeRanges);
assert.equal(tableCloneFixture.verify().ok, true);
const tableCloneSourcePptx = await PresentationFile.exportPptx(tableCloneFixture);
const tableCloneSourceZip = await JSZip.loadAsync(tableCloneSourcePptx.bytes);
const tableCloneSourceXml = await tableCloneSourceZip.file("ppt/slides/slide1.xml").async("text");
assert.match(tableCloneSourceXml, /<a:tc gridSpan="3">/);
assert.match(tableCloneSourceXml, /<a:tc hMerge="1">/);
assert.match(tableCloneSourceXml, /<a:tc rowSpan="2" gridSpan="2">/);
assert.match(tableCloneSourceXml, /<a:tc hMerge="1" vMerge="1">/);
const tableCloneImportedDeck = await PresentationFile.importPptx(tableCloneSourcePptx);
const tableCloneImportedSource = tableCloneImportedDeck.slides.getItem(0);
assert.deepEqual(tableCloneImportedSource.tables.items[0].mergeRanges, tableCloneSourceTable.mergeRanges);
const tableClone = tableCloneImportedSource.duplicate();
assert.equal(tableClone.tables.items.length, 1);
assert.notEqual(tableClone.tables.items[0], tableCloneImportedSource.tables.items[0]);
assert.notEqual(tableClone.tables.items[0].id, tableCloneImportedSource.tables.items[0].id);
assert.deepEqual(tableClone.tables.items[0].values, tableCloneSourceTable.values);
assert.deepEqual(tableClone.tables.items[0].mergeRanges, tableCloneSourceTable.mergeRanges);
const tableClonePptx = await PresentationFile.exportPptx(tableCloneImportedDeck);
const tableCloneZip = await JSZip.loadAsync(tableClonePptx.bytes);
const tableCloneSlidePath = (await orderedPptxSlidePaths(tableCloneZip))[1];
assert.deepEqual(
  await tableCloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await tableCloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning an inline table must retain the origin SlidePart byte-for-byte",
);
assert.ok(tableCloneZip.file(tableCloneSlidePath), "the table clone must own a new SlidePart");
const tableCloneRelationships = await tableCloneZip.file(relationshipPartPath(tableCloneSlidePath)).async("text");
assert.match(tableCloneRelationships, /\/slideLayout/);
assert.doesNotMatch(tableCloneRelationships, /\/(?:image|chart|hyperlink|oleObject|package)"/i, "canonical table clones must not add a table-specific OPC edge");
const tableCloneRoundTrip = await PresentationFile.importPptx(tableClonePptx);
assert.deepEqual(tableCloneRoundTrip.slides.items.map((slide) => slide.tables.items[0].values), [
  tableCloneSourceTable.values,
  tableCloneSourceTable.values,
]);
assert.deepEqual(tableCloneRoundTrip.slides.items.map((slide) => slide.tables.items[0].mergeRanges), [tableCloneSourceTable.mergeRanges, tableCloneSourceTable.mergeRanges]);
tableCloneRoundTrip.slides.getItem(1).tables.items[0].cells.set(2, 2, "Edited after table clone reimport");
const tableCloneEditedPptx = await PresentationFile.exportPptx(tableCloneRoundTrip);
const tableCloneEditedRoundTrip = await PresentationFile.importPptx(tableCloneEditedPptx);
assert.equal(tableCloneEditedRoundTrip.slides.getItem(1).tables.items[0].values[2][2], "Edited after table clone reimport");
assert.equal(tableCloneEditedRoundTrip.slides.getItem(0).tables.items[0].values[2][2], "Release");

const immediateTableCloneEdit = await PresentationFile.importPptx(tableCloneSourcePptx);
immediateTableCloneEdit.slides.getItem(0).duplicate().tables.items[0].values[2][2] = "Too soon";
await assert.rejects(
  () => PresentationFile.exportPptx(immediateTableCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

const importedTableMergeChange = await PresentationFile.importPptx(tableCloneSourcePptx);
importedTableMergeChange.slides.getItem(0).tables.items[0].merge({ startRow: 1, endRow: 2, startColumn: 2, endColumn: 2 });
await assert.rejects(
  () => PresentationFile.exportPptx(importedTableMergeChange),
  (error) => error?.code === "unsupported_presentation_edit",
);

// A recognized literal-data chart may travel with the bounded slide clone
// only when its ChartPart has no child/external/data relationship graph. The
// clone copies the exact chart XML into a distinct part: sharing would make a
// later chart edit affect both slides.
const chartCloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const chartCloneOriginal = chartCloneFixture.slides.add({ name: "Chart original" });
chartCloneOriginal.charts.add("bar", {
  name: "pipeline-chart",
  position: { left: 48, top: 42, width: 500, height: 250 },
  title: "Quarterly pipeline",
  categories: ["Q1", "Q2", "Q3"],
  series: [{ name: "Pipeline", values: [42, 48, 57], fill: "#2563EB" }],
  axes: {
    category: { title: "Quarter" },
    value: { title: "Value", min: 0, max: 80, majorUnit: 20 },
  },
  legend: { visible: true, position: "r" },
  dataLabels: { showValue: true, position: "top" },
});
const chartCloneSourcePptx = await PresentationFile.exportPptx(chartCloneFixture);
const chartCloneSourceZip = await JSZip.loadAsync(chartCloneSourcePptx.bytes);
const chartPartPaths = (zip) => Object.keys(zip.files)
  .filter((partPath) => /^ppt\/(?:slides\/)?charts\/chart\d+\.xml$/i.test(partPath))
  .sort();
const chartCloneSourceParts = chartPartPaths(chartCloneSourceZip);
assert.equal(chartCloneSourceParts.length, 1, "the source fixture must own exactly one ChartPart");
const [chartCloneSourcePart] = chartCloneSourceParts;
const chartCloneImportedDeck = await PresentationFile.importPptx(chartCloneSourcePptx);
const chartCloneImportedSource = chartCloneImportedDeck.slides.getItem(0);
const chartClonePending = chartCloneImportedSource.duplicate();
assert.equal(chartClonePending.charts.items.length, 1);
assert.notEqual(chartClonePending.charts.items[0], chartCloneImportedSource.charts.items[0]);
assert.notEqual(chartClonePending.charts.items[0].id, chartCloneImportedSource.charts.items[0].id);
assert.deepEqual(chartClonePending.charts.items[0].categories, ["Q1", "Q2", "Q3"]);
assert.deepEqual(chartClonePending.charts.items[0].series[0].values, [42, 48, 57]);
const chartClonePptx = await PresentationFile.exportPptx(chartCloneImportedDeck);
const chartCloneZip = await JSZip.loadAsync(chartClonePptx.bytes);
const chartCloneSlidePath = (await orderedPptxSlidePaths(chartCloneZip))[1];
assert.deepEqual(
  await chartCloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await chartCloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning a closed native chart must retain the origin SlidePart byte-for-byte",
);
const chartCloneOutputParts = chartPartPaths(chartCloneZip);
assert.equal(chartCloneOutputParts.length, 2, "the chart clone must allocate exactly one additional ChartPart");
const chartCloneNewParts = chartCloneOutputParts.filter((partPath) => !chartCloneSourceZip.file(partPath));
assert.equal(chartCloneNewParts.length, 1, "the chart clone must own one distinct new ChartPart path");
const [chartClonePart] = chartCloneNewParts;
assert.deepEqual(
  await chartCloneZip.file(chartClonePart).async("uint8array"),
  await chartCloneSourceZip.file(chartCloneSourcePart).async("uint8array"),
  "the first clone export must copy the accepted ChartPart byte-for-byte",
);
const modeledChartRelationship = (xml) => {
  const relationships = [...xml.matchAll(/<Relationship\b[^>]*>/g)]
    .map(([tag]) => ({
      id: /\bId="([^"]+)"/.exec(tag)?.[1],
      type: /\bType="([^"]+)"/.exec(tag)?.[1],
      target: /\bTarget="([^"]+)"/.exec(tag)?.[1],
      targetMode: /\bTargetMode="([^"]+)"/.exec(tag)?.[1],
    }))
    .filter((relationship) => /\/chart$/.test(relationship.type || ""));
  assert.equal(relationships.length, 1, "a bounded chart clone slide must own exactly one chart relationship");
  return relationships[0];
};
const sourceChartRelationship = modeledChartRelationship(await chartCloneZip.file("ppt/slides/_rels/slide1.xml.rels").async("text"));
const cloneChartRelationship = modeledChartRelationship(await chartCloneZip.file(relationshipPartPath(chartCloneSlidePath)).async("text"));
assert.equal(cloneChartRelationship.id, sourceChartRelationship.id, "the clone must retain the slide-local chart relationship ID");
assert.equal(cloneChartRelationship.type, sourceChartRelationship.type);
assert.equal(cloneChartRelationship.targetMode, undefined);
const resolvedCloneChartTarget = cloneChartRelationship.target.startsWith("/")
  ? cloneChartRelationship.target.replace(/^\/+/, "")
  : path.posix.normalize(path.posix.join("ppt/slides", cloneChartRelationship.target));
assert.equal(resolvedCloneChartTarget, chartClonePart, "the clone chart relationship must target its newly allocated ChartPart");

const chartCloneRoundTrip = await PresentationFile.importPptx(chartClonePptx);
assert.deepEqual(chartCloneRoundTrip.slides.items.map((slide) => slide.charts.items[0].title), ["Quarterly pipeline", "Quarterly pipeline"]);
chartCloneRoundTrip.slides.getItem(1).charts.items[0].title = "Updated clone pipeline";
chartCloneRoundTrip.slides.getItem(1).charts.items[0].series[0].values[1] = 63;
const chartCloneEditedPptx = await PresentationFile.exportPptx(chartCloneRoundTrip);
const chartCloneEditedZip = await JSZip.loadAsync(chartCloneEditedPptx.bytes);
assert.deepEqual(
  await chartCloneEditedZip.file(chartCloneSourcePart).async("uint8array"),
  await chartCloneSourceZip.file(chartCloneSourcePart).async("uint8array"),
  "editing the reimported clone chart must leave the origin ChartPart byte-for-byte intact",
);
const chartCloneEditedRoundTrip = await PresentationFile.importPptx(chartCloneEditedPptx);
assert.equal(chartCloneEditedRoundTrip.slides.getItem(0).charts.items[0].title, "Quarterly pipeline");
assert.equal(chartCloneEditedRoundTrip.slides.getItem(0).charts.items[0].series[0].values[1], 48);
assert.equal(chartCloneEditedRoundTrip.slides.getItem(1).charts.items[0].title, "Updated clone pipeline");
assert.equal(chartCloneEditedRoundTrip.slides.getItem(1).charts.items[0].series[0].values[1], 63);

const immediateChartCloneEdit = await PresentationFile.importPptx(chartCloneSourcePptx);
immediateChartCloneEdit.slides.getItem(0).duplicate().charts.items[0].title = "Too soon";
await assert.rejects(
  () => PresentationFile.exportPptx(immediateChartCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

const connectedChartCloneZip = await JSZip.loadAsync(chartCloneSourcePptx.bytes);
const chartRelationshipPart = path.posix.join(
  path.posix.dirname(chartCloneSourcePart),
  "_rels",
  path.posix.basename(chartCloneSourcePart) + ".rels",
);
connectedChartCloneZip.file(
  chartRelationshipPart,
  '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdUnsafeChartChild" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/chart-child" TargetMode="External"/></Relationships>',
);
const connectedChartCloneSource = await connectedChartCloneZip.generateAsync({ type: "uint8array" });
const connectedChartCloneDeck = await PresentationFile.importPptx(connectedChartCloneSource);
connectedChartCloneDeck.slides.getItem(0).duplicate();
const connectedChartCloneOutput = await PresentationFile.exportPptx(connectedChartCloneDeck);
const connectedChartCloneOutputZip = await JSZip.loadAsync(connectedChartCloneOutput.bytes);
const connectedChartParts = chartPartPaths(connectedChartCloneOutputZip);
assert.equal(connectedChartParts.length, 2);
for (const chartPartPath of connectedChartParts) {
  assert.match(
    await connectedChartCloneOutputZip.file(relationshipPartPath(chartPartPath)).async("text"),
    /rIdUnsafeChartChild/,
    "the graph clone must retain external relationships owned by both independent ChartParts",
  );
}

// A group is clone-safe only when every descendant is already in the narrow
// shape/table/closed-chart/image leaf. The group and every child receive fresh
// JS identity, while nested relationship-owning leaves still require the same
// native preflight as their top-level counterparts.
const groupCloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const groupCloneOriginal = groupCloneFixture.slides.add({ name: "Recursive group original" });
const groupCloneRoot = groupCloneOriginal.addGroup({
  name: "evidence-cluster",
  position: { left: 48, top: 40, width: 420, height: 220 },
  childFrame: { left: 0, top: 0, width: 420, height: 220 },
});
groupCloneRoot.shapes.add({
  name: "group-copy",
  position: { left: 0, top: 0, width: 220, height: 48 },
  text: "Clone-safe grouped evidence",
  fill: "#FFFFFF",
  line: { fill: "#2563EB", width: 1 },
});
groupCloneRoot.tables.add({
  name: "group-decision-grid",
  position: { left: 0, top: 66, width: 220, height: 110 },
  values: [["Gate", "State"], ["QA", "Pass"]],
  rows: 2,
  columns: 2,
  styleOptions: { headerRow: true, bandedRows: true },
});
groupCloneRoot.images.add({
  name: "group-immutable-asset",
  alt: "Shared nested clone asset",
  position: { left: 246, top: 0, width: 96, height: 96 },
  dataUrl: PNG,
  fit: "stretch",
});
const groupCloneNested = groupCloneRoot.groups.add({
  name: "nested-evidence",
  position: { left: 246, top: 118, width: 150, height: 48 },
  childFrame: { left: 0, top: 0, width: 150, height: 48 },
});
groupCloneNested.shapes.add({
  name: "nested-copy",
  position: { left: 0, top: 0, width: 150, height: 48 },
  text: "Nested",
  fill: "#DBEAFE",
  line: { fill: "#2563EB", width: 1 },
});
const groupCloneSourcePptx = await PresentationFile.exportPptx(groupCloneFixture);
const groupCloneSourceZip = await JSZip.loadAsync(groupCloneSourcePptx.bytes);
const groupCloneSourceMediaPaths = Object.keys(groupCloneSourceZip.files)
  .filter((path) => /^ppt\/media\/[^/]+\.(?:png|jpe?g|gif|svg)$/i.test(path));
assert.equal(groupCloneSourceMediaPaths.length, 1, "the recursive group fixture must contain one nested image part");
const [groupCloneSourceMediaPath] = groupCloneSourceMediaPaths;
const groupCloneImportedDeck = await PresentationFile.importPptx(groupCloneSourcePptx);
const groupCloneImportedSource = groupCloneImportedDeck.slides.getItem(0);
const groupClone = groupCloneImportedSource.duplicate();
const groupCloneCopy = groupClone.groups.items[0];
assert.equal(groupClone.groups.items.length, 1);
assert.notEqual(groupCloneCopy, groupCloneImportedSource.groups.items[0]);
assert.notEqual(groupCloneCopy.id, groupCloneImportedSource.groups.items[0].id);
assert.notEqual(groupCloneCopy.shapes.items[0].id, groupCloneImportedSource.groups.items[0].shapes.items[0].id);
assert.notEqual(groupCloneCopy.tables.items[0].id, groupCloneImportedSource.groups.items[0].tables.items[0].id);
assert.notEqual(groupCloneCopy.images.items[0].id, groupCloneImportedSource.groups.items[0].images.items[0].id);
assert.notEqual(groupCloneCopy.groups.items[0].id, groupCloneImportedSource.groups.items[0].groups.items[0].id);
assert.deepEqual(groupCloneCopy.tables.items[0].values, [["Gate", "State"], ["QA", "Pass"]]);
assert.equal(groupCloneCopy.groups.items[0].shapes.items[0].text.value, "Nested");
const groupClonePptx = await PresentationFile.exportPptx(groupCloneImportedDeck);
const groupCloneZip = await JSZip.loadAsync(groupClonePptx.bytes);
const groupCloneSlidePath = (await orderedPptxSlidePaths(groupCloneZip))[1];
assert.deepEqual(
  await groupCloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await groupCloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning a recursively canonical group must retain its origin SlidePart byte-for-byte",
);
assert.deepEqual(
  await groupCloneZip.file(groupCloneSourceMediaPath).async("uint8array"),
  await groupCloneSourceZip.file(groupCloneSourceMediaPath).async("uint8array"),
  "cloning a recursively canonical group must retain its shared media bytes",
);
assert.deepEqual(
  imageRelationshipTargets(await groupCloneZip.file(relationshipPartPath(groupCloneSlidePath)).async("text")),
  imageRelationshipTargets(await groupCloneZip.file("ppt/slides/_rels/slide1.xml.rels").async("text")),
  "nested clone images must receive the same verified relationship target as the origin",
);
const groupCloneRoundTrip = await PresentationFile.importPptx(groupClonePptx);
assert.equal(groupCloneRoundTrip.slides.items.length, 2);
assert.equal(groupCloneRoundTrip.slides.getItem(1).groups.items[0].groups.items[0].shapes.items[0].text.value, "Nested");
groupCloneRoundTrip.slides.getItem(1).groups.items[0].groups.items[0].shapes.items[0].text.set("Edited after group clone reimport");
const groupCloneEditedPptx = await PresentationFile.exportPptx(groupCloneRoundTrip);
const groupCloneEditedRoundTrip = await PresentationFile.importPptx(groupCloneEditedPptx);
assert.equal(groupCloneEditedRoundTrip.slides.getItem(1).groups.items[0].groups.items[0].shapes.items[0].text.value, "Edited after group clone reimport");

const immediateGroupCloneEdit = await PresentationFile.importPptx(groupCloneSourcePptx);
immediateGroupCloneEdit.slides.getItem(0).duplicate().groups.items[0].groups.items[0].shapes.items[0].text.set("Too soon");
await assert.rejects(
  () => PresentationFile.exportPptx(immediateGroupCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

const connectedGroupCloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const connectedGroupCloneOriginal = connectedGroupCloneFixture.slides.add({ name: "Connected group original" });
const connectedGroupCloneRoot = connectedGroupCloneOriginal.addGroup({
  name: "connected-cluster",
  position: { left: 48, top: 40, width: 320, height: 120 },
  childFrame: { left: 0, top: 0, width: 320, height: 120 },
});
const connectedGroupCloneLeft = connectedGroupCloneRoot.shapes.add({ name: "left", position: { left: 0, top: 20, width: 90, height: 42 }, text: "Left" });
const connectedGroupCloneRight = connectedGroupCloneRoot.shapes.add({ name: "right", position: { left: 210, top: 20, width: 90, height: 42 }, text: "Right" });
connectedGroupCloneRoot.connectors.add({
  name: "join",
  from: connectedGroupCloneLeft,
  to: connectedGroupCloneRight,
  start: { x: 90, y: 41 },
  end: { x: 210, y: 41 },
  line: { fill: "#64748B", width: 1 },
});
const connectedGroupCloneSourcePptx = await PresentationFile.exportPptx(connectedGroupCloneFixture);
const connectedGroupCloneSourceZip = await JSZip.loadAsync(connectedGroupCloneSourcePptx.bytes);
const connectedGroupCloneImported = await PresentationFile.importPptx(connectedGroupCloneSourcePptx);
const connectedGroupCloneImportedSource = connectedGroupCloneImported.slides.getItem(0);
const connectedGroupCloneSourceGroup = connectedGroupCloneImportedSource.groups.items[0];
const connectedGroupCloneSourceConnector = connectedGroupCloneSourceGroup.connectors.items[0];
assert.equal(connectedGroupCloneSourceConnector.startTargetId, connectedGroupCloneSourceGroup.shapes.items[0].id);
assert.equal(connectedGroupCloneSourceConnector.endTargetId, connectedGroupCloneSourceGroup.shapes.items[1].id);
const connectedGroupClone = connectedGroupCloneImportedSource.duplicate();
const connectedGroupCloneCopy = connectedGroupClone.groups.items[0];
const connectedGroupCloneConnector = connectedGroupCloneCopy.connectors.items[0];
assert.notEqual(connectedGroupCloneConnector.id, connectedGroupCloneSourceConnector.id);
assert.notEqual(connectedGroupCloneCopy.shapes.items[0].id, connectedGroupCloneSourceGroup.shapes.items[0].id);
assert.equal(connectedGroupCloneConnector.startTargetId, connectedGroupCloneCopy.shapes.items[0].id);
assert.equal(connectedGroupCloneConnector.endTargetId, connectedGroupCloneCopy.shapes.items[1].id);
const connectedGroupClonePptx = await PresentationFile.exportPptx(connectedGroupCloneImported);
const connectedGroupCloneZip = await JSZip.loadAsync(connectedGroupClonePptx.bytes);
const connectedGroupCloneSlidePath = (await orderedPptxSlidePaths(connectedGroupCloneZip))[1];
assert.deepEqual(
  await connectedGroupCloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await connectedGroupCloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning a group with bounded connectors must retain its origin SlidePart byte-for-byte",
);
assert.ok(connectedGroupCloneZip.file(connectedGroupCloneSlidePath), "the connector clone must own a new SlidePart");
const connectedGroupCloneRoundTrip = await PresentationFile.importPptx(connectedGroupClonePptx);
const connectedGroupCloneRoundTripGroup = connectedGroupCloneRoundTrip.slides.getItem(1).groups.items[0];
const connectedGroupCloneRoundTripConnector = connectedGroupCloneRoundTripGroup.connectors.items[0];
assert.equal(connectedGroupCloneRoundTripConnector.startTargetId, connectedGroupCloneRoundTripGroup.shapes.items[0].id);
assert.equal(connectedGroupCloneRoundTripConnector.endTargetId, connectedGroupCloneRoundTripGroup.shapes.items[1].id);

const immediateConnectedGroupCloneEdit = await PresentationFile.importPptx(connectedGroupCloneSourcePptx);
immediateConnectedGroupCloneEdit.slides.getItem(0).duplicate().groups.items[0].connectors.items[0].line.width = 2;
await assert.rejects(
  () => PresentationFile.exportPptx(immediateConnectedGroupCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

const unresolvedConnectedGroupClone = await PresentationFile.importPptx(connectedGroupCloneSourcePptx);
unresolvedConnectedGroupClone.slides.getItem(0).groups.items[0].connectors.items[0].startTargetId = "missing-source-target";
assert.throws(
  () => unresolvedConnectedGroupClone.slides.getItem(0).duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);
assert.equal(unresolvedConnectedGroupClone.slides.items.length, 1, "connector-target preflight must not leave a partial clone behind");

// Speaker notes add one deliberately closed relationship leaf to the same
// clone profile. The NotesSlide itself is new and points at the clone, while
// its NotesMaster stays immutable and shared. This is raw part preservation,
// not permission to edit notes before the export/reimport boundary.
const notesCloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const notesCloneOriginal = notesCloneFixture.slides.add({
  name: "Notes image original",
  notes: "Open with the customer outcome.\nClose with the operating decision.",
});
notesCloneOriginal.shapes.add({ name: "notes-clone-copy", position: { left: 48, top: 48, width: 300, height: 72 }, text: "Notes image original" });
notesCloneOriginal.images.add({
  name: "notes-clone-asset",
  alt: "Notes clone immutable asset",
  position: { left: 48, top: 150, width: 120, height: 120 },
  dataUrl: PNG,
  fit: "stretch",
});
const notesCloneSourcePptx = await PresentationFile.exportPptx(notesCloneFixture);
const notesCloneSourceZip = await JSZip.loadAsync(notesCloneSourcePptx.bytes);
const notesCloneSourceNotesPaths = Object.keys(notesCloneSourceZip.files)
  .filter((path) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(path));
assert.deepEqual(notesCloneSourceNotesPaths, ["ppt/notesSlides/notesSlide1.xml"]);
const [notesCloneSourceNotesPath] = notesCloneSourceNotesPaths;
const notesCloneSourceMediaPath = Object.keys(notesCloneSourceZip.files)
  .find((path) => /^ppt\/media\/[^/]+\.(?:png|jpe?g|gif|svg)$/i.test(path));
assert.ok(notesCloneSourceMediaPath, "the notes clone fixture must contain one embedded image part");
const notesCloneImportedDeck = await PresentationFile.importPptx(notesCloneSourcePptx);
const notesCloneImportedSource = notesCloneImportedDeck.slides.getItem(0);
const notesClone = notesCloneImportedSource.duplicate();
assert.equal(notesClone.speakerNotes.text, "Open with the customer outcome.\nClose with the operating decision.");
const notesClonePptx = await PresentationFile.exportPptx(notesCloneImportedDeck);
const notesCloneZip = await JSZip.loadAsync(notesClonePptx.bytes);
const notesCloneSlidePath = (await orderedPptxSlidePaths(notesCloneZip))[1];
assert.deepEqual(
  await notesCloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await notesCloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning a slide with notes must retain the origin SlidePart byte-for-byte",
);
assert.deepEqual(
  await notesCloneZip.file(notesCloneSourceNotesPath).async("uint8array"),
  await notesCloneSourceZip.file(notesCloneSourceNotesPath).async("uint8array"),
  "cloning a slide with notes must retain the origin NotesSlide byte-for-byte",
);
assert.deepEqual(
  await notesCloneZip.file(notesCloneSourceMediaPath).async("uint8array"),
  await notesCloneSourceZip.file(notesCloneSourceMediaPath).async("uint8array"),
  "cloning a slide with notes must retain the shared media bytes",
);
const notesClonePaths = Object.keys(notesCloneZip.files)
  .filter((path) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(path));
assert.equal(notesClonePaths.length, 2);
const notesCloneCopyPath = notesClonePaths.find((partPath) => partPath !== notesCloneSourceNotesPath);
assert.ok(notesCloneCopyPath);
assert.equal(
  await notesCloneZip.file(notesCloneCopyPath).async("text"),
  await notesCloneZip.file(notesCloneSourceNotesPath).async("text"),
  "the clone NotesSlide XML must be a verbatim copy of the source notes XML",
);
const relationshipTagForType = (relationships, suffix) => [...relationships.matchAll(/<Relationship\b[^>]*>/gi)]
  .find(([tag]) => new RegExp(`\\bType="[^"]*\\/${suffix}"`, "i").test(tag))?.[0];
const relationshipAttributeForType = (relationships, suffix, attribute) => {
  const tag = relationshipTagForType(relationships, suffix);
  return tag && new RegExp(`\\b${attribute}="([^"]+)"`, "i").exec(tag)?.[1];
};
const sourceSlideRelationships = await notesCloneZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const cloneSlideRelationships = await notesCloneZip.file(relationshipPartPath(notesCloneSlidePath)).async("text");
const sourceNotesRelationships = await notesCloneZip.file("ppt/notesSlides/_rels/notesSlide1.xml.rels").async("text");
const cloneNotesRelationships = await notesCloneZip.file(relationshipPartPath(notesCloneCopyPath)).async("text");
assert.equal(relationshipAttributeForType(cloneSlideRelationships, "notesSlide", "Id"), relationshipAttributeForType(sourceSlideRelationships, "notesSlide", "Id"));
assert.equal(relationshipAttributeForType(cloneNotesRelationships, "notesMaster", "Id"), relationshipAttributeForType(sourceNotesRelationships, "notesMaster", "Id"));
assert.equal(relationshipAttributeForType(cloneNotesRelationships, "notesMaster", "Target"), relationshipAttributeForType(sourceNotesRelationships, "notesMaster", "Target"));
assert.equal(relationshipAttributeForType(cloneNotesRelationships, "slide", "Id"), relationshipAttributeForType(sourceNotesRelationships, "slide", "Id"));
assert.equal(relationshipAttributeForType(cloneNotesRelationships, "slide", "Target"), `/${notesCloneSlidePath}`);
const notesCloneRoundTrip = await PresentationFile.importPptx(notesClonePptx);
assert.deepEqual(notesCloneRoundTrip.slides.items.map((slide) => slide.speakerNotes.text), [
  "Open with the customer outcome.\nClose with the operating decision.",
  "Open with the customer outcome.\nClose with the operating decision.",
]);
notesCloneRoundTrip.slides.getItem(1).speakerNotes.text = "Edited after notes clone reimport.";
const notesCloneEditedPptx = await PresentationFile.exportPptx(notesCloneRoundTrip);
const notesCloneEditedRoundTrip = await PresentationFile.importPptx(notesCloneEditedPptx);
assert.equal(notesCloneEditedRoundTrip.slides.getItem(1).speakerNotes.text, "Edited after notes clone reimport.");

const immediateNotesCloneEdit = await PresentationFile.importPptx(notesCloneSourcePptx);
immediateNotesCloneEdit.slides.getItem(0).duplicate().speakerNotes.text = "Too soon";
await assert.rejects(
  () => PresentationFile.exportPptx(immediateNotesCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

// Deletion is graph-ownership based rather than object-type based. Pick one
// imported slide with a nontrivial exclusive descendant closure and prove the
// codec removes it without asking JavaScript to rebuild its chart/media graph.
const complexImportedDeletion = await PresentationFile.importPptx(firstExport);
const complexDeletableSlide = complexImportedDeletion.slides.items.find((slide) =>
  slide.deletionCapability.supported && slide.deletionCapability.ownedPartCount > 1);
assert.ok(complexDeletableSlide, "fixture must expose one slide-owned nontrivial OPC closure");
const complexDeletedName = complexDeletableSlide.name;
const complexSlideCountBeforeDelete = complexImportedDeletion.slides.count;
complexDeletableSlide.delete();
const complexDeletionPptx = await PresentationFile.exportPptx(complexImportedDeletion);
const complexDeletionRoundTrip = await PresentationFile.importPptx(complexDeletionPptx);
assert.equal(complexDeletionRoundTrip.slides.count, complexSlideCountBeforeDelete - 1);
assert.ok(!complexDeletionRoundTrip.slides.items.some((slide) => slide.name === complexDeletedName));

// Turn the canonical package into a source-bound template marker without
// creating a second writer, then prove its Master/Layout parts survive a
// modeled slide edit byte-for-byte.
const firstZip = await JSZip.loadAsync(new Uint8Array(await firstExport.arrayBuffer()));
const firstSlideXml = await firstZip.file("ppt/slides/slide1.xml").async("text");
assert.match(firstSlideXml, /<a:srcRect[^>]*l="25000"/);
assert.match(firstSlideXml, /<a:srcRect[^>]*r="25000"/);
const authoredChartEntries = await Promise.all(Object.keys(firstZip.files)
  .filter((name) => /\/charts\/chart\d+\.xml$/.test(name))
  .map(async (name) => ({ name, xml: await firstZip.file(name).async("text") })));
const authoredChartXml = authoredChartEntries.map((entry) => entry.xml);
const lineChartEntry = authoredChartEntries.find((entry) => entry.xml.includes("Trend"));
assert.ok(lineChartEntry);
assert.equal((lineChartEntry.xml.match(/<c:trendline>/g) || []).length, 3);
assert.deepEqual([...lineChartEntry.xml.matchAll(/<c:trendlineType val="([^"]+)"\s*\/>/g)].map((match) => match[1]), ["linear", "movingAvg", "poly"]);
assert.match(lineChartEntry.xml, /<c:forward val="0\.5"\s*\/>/);
assert.match(lineChartEntry.xml, /<c:dispEq val="1"\s*\/>/);
assert.match(lineChartEntry.xml, /<c:errValType val="stdDev"\s*\/>[\s\S]*?<c:val val="1\.5"\s*\/>/);
assert.match(lineChartEntry.xml, /<c:noEndCap val="1"\s*\/>/);
const comboChartXml = authoredChartXml.find((xml) => xml.includes("Revenue and margin"));
assert.ok(comboChartXml);
assert.match(comboChartXml, /<c:barChart>/);
assert.match(comboChartXml, /<c:lineChart>/);
assert.match(comboChartXml, /<c:trendlineType val="exp"\s*\/>/);
assert.match(comboChartXml, /<c:errBarType val="minus"\s*\/>[\s\S]*?<c:errValType val="cust"\s*\/>[\s\S]*?<c:minus><c:numLit>/);
assert.match(comboChartXml, /<c:barChart>[\s\S]*?<c:axId val="1"\s*\/><c:axId val="2"\s*\/><\/c:barChart>/);
assert.match(comboChartXml, /<c:lineChart>[\s\S]*?<c:axId val="1"\s*\/><c:axId val="2"\s*\/><\/c:lineChart>/);

const unsupportedTrendlineLabelXml = lineChartEntry.xml.replace("</c:trendline>", "<c:trendlineLbl/></c:trendline>");
assert.notEqual(unsupportedTrendlineLabelXml, lineChartEntry.xml);
const unsupportedTrendlineLabelSource = await PresentationFile.patchPptx(firstExport, [
  { path: lineChartEntry.name, xml: unsupportedTrendlineLabelXml },
]);
const preservedTrendlineLabelDeck = await PresentationFile.importPptx(unsupportedTrendlineLabelSource);
const preservedTrendlineLabelChart = itemByName(preservedTrendlineLabelDeck.slides.getItem(1).charts.items, "line-chart");
assert.equal(preservedTrendlineLabelChart.series[0].trendlines, undefined);
const preservedTrendlineLabelOutput = await PresentationFile.exportPptx(preservedTrendlineLabelDeck);
const preservedTrendlineLabelZip = await JSZip.loadAsync(preservedTrendlineLabelOutput.bytes);
assert.equal(await preservedTrendlineLabelZip.file(lineChartEntry.name).async("text"), unsupportedTrendlineLabelXml);
preservedTrendlineLabelChart.title = "Forbidden trendline-label edit";
await assert.rejects(
  () => PresentationFile.exportPptx(preservedTrendlineLabelDeck),
  (error) => error?.code === "unsupported_presentation_edit" && /preserved but not safely editable/i.test(error.message),
);

const unsupportedErrorBarsXml = lineChartEntry.xml.replace("</c:errBars>", "<c:extLst/></c:errBars>");
assert.notEqual(unsupportedErrorBarsXml, lineChartEntry.xml);
const unsupportedErrorBarsSource = await PresentationFile.patchPptx(firstExport, [
  { path: lineChartEntry.name, xml: unsupportedErrorBarsXml },
]);
const preservedErrorBarsDeck = await PresentationFile.importPptx(unsupportedErrorBarsSource);
const preservedErrorBarsChart = itemByName(preservedErrorBarsDeck.slides.getItem(1).charts.items, "line-chart");
assert.equal(preservedErrorBarsChart.series[0].errorBars, undefined);
const preservedErrorBarsOutput = await PresentationFile.exportPptx(preservedErrorBarsDeck);
const preservedErrorBarsZip = await JSZip.loadAsync(preservedErrorBarsOutput.bytes);
assert.equal(await preservedErrorBarsZip.file(lineChartEntry.name).async("text"), unsupportedErrorBarsXml);
preservedErrorBarsChart.title = "Forbidden error-bar extension edit";
await assert.rejects(
  () => PresentationFile.exportPptx(preservedErrorBarsDeck),
  (error) => error?.code === "unsupported_presentation_edit" && /preserved but not safely editable/i.test(error.message),
);

// Reference 2.8.24 exposes imported PowerPoint grid spacing, snap settings,
// and guides through presentation.view. The project retains those local
// projections while adding only the separately re-proven fixed-topology edit.
const viewPropertiesXml = '<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" lastView="sldView"><p:slideViewPr><p:cSldViewPr snapToGrid="1" snapToObjects="0" showGuides="1"><p:cViewPr varScale="1"><p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale><p:origin x="0" y="0"/></p:cViewPr><p:guideLst><p:guide orient="horz" pos="2160"/><p:guide orient="vert" pos="2880"/></p:guideLst></p:cSldViewPr></p:slideViewPr><p:gridSpacing cx="72008" cy="91440"/></p:viewPr>';
const presentationRelationships = await firstZip.file("ppt/_rels/presentation.xml.rels").async("text");
const viewSource = await PresentationFile.patchPptx(firstExport, [
  {
    path: "ppt/_rels/presentation.xml.rels",
    xml: presentationRelationships.replace("</Relationships>", '<Relationship Id="rIdViewProperties" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/></Relationships>'),
  },
  {
    path: "ppt/viewProps.xml",
    xml: viewPropertiesXml,
    contentType: "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml",
  },
]);
const importedViewPresentation = await PresentationFile.importPptx(viewSource);
assert.equal(importedViewPresentation.view.gridSpacingCxEmu, 72_008);
assert.equal(importedViewPresentation.view.gridSpacingCyEmu, 91_440);
assert.equal(importedViewPresentation.view.gridlinesVisible, false);
assert.equal(importedViewPresentation.view.guidesVisible, false);
assert.deepEqual(importedViewPresentation.view.capability, {
  sourceBound: true,
  partPresent: true,
  editable: true,
  gridSpacingCxEmuPresent: true,
  gridSpacingCyEmuPresent: true,
  slideViewSnapToGridPresent: true,
  slideViewSnapToObjectsPresent: true,
  guideCount: 2,
});
assert.deepEqual(importedViewPresentation.view.toProto(), {
  gridSpacingCxEmu: 72_008,
  gridSpacingCyEmu: 91_440,
  slideViewSnapToGrid: true,
  slideViewSnapToObjects: false,
  slideViewShowGuides: false,
  slideGuides: [
    { orientation: "horizontal", position: 2160 },
    { orientation: "vertical", position: 2880 },
  ],
});
assert.deepEqual(importedViewPresentation.master.slideGuides, importedViewPresentation.view.toProto().slideGuides);
assert.deepEqual(importedViewPresentation.layouts.items[0].slideGuides, importedViewPresentation.view.toProto().slideGuides);
assert.throws(() => importedViewPresentation.layouts.items[0].slideGuides[0].position = 0, TypeError);
assert.equal(importedViewPresentation.view.showGridlines(), undefined);
assert.equal(importedViewPresentation.view.showGuides(), undefined);
assert.equal(importedViewPresentation.view.gridlinesVisible, true);
assert.equal(importedViewPresentation.view.guidesVisible, true);
const viewRoundTripFile = await PresentationFile.exportPptx(importedViewPresentation);
const viewRoundTripZip = await JSZip.loadAsync(viewRoundTripFile.bytes);
assert.equal(await viewRoundTripZip.file("ppt/viewProps.xml").async("text"), viewPropertiesXml);
const viewRoundTrip = await PresentationFile.importPptx(viewRoundTripFile);
assert.deepEqual(viewRoundTrip.view.toProto().slideGuides, importedViewPresentation.view.toProto().slideGuides);
assert.throws(
  () => viewRoundTrip.view.setSourceProperties({ slideViewShowGuides: true }),
  /unsupported fields/,
);
assert.throws(
  () => viewRoundTrip.view.setSourceProperties({
    slideGuides: [
      { orientation: "vertical", position: 2161 },
      { orientation: "horizontal", position: 2881 },
    ],
  }),
  /guide count, order, and orientation are source-bound/,
);
viewRoundTrip.view.setSourceProperties({
  gridSpacingCxEmu: 72_009,
  gridSpacingCyEmu: 91_441,
  slideViewSnapToGrid: false,
  slideViewSnapToObjects: true,
  slideGuides: [
    { orientation: "horizontal", position: 2161 },
    { orientation: "vertical", position: 2881 },
  ],
});
const viewEditedFile = await PresentationFile.exportPptx(viewRoundTrip);
const viewEditedZip = await JSZip.loadAsync(viewEditedFile.bytes);
assert.deepEqual(Object.keys(viewEditedZip.files).sort(), Object.keys(viewRoundTripZip.files).sort());
for (const partPath of Object.keys(viewRoundTripZip.files).filter((name) => !viewRoundTripZip.files[name].dir && name !== "ppt/viewProps.xml")) {
  assert.deepEqual(
    await viewEditedZip.file(partPath).async("uint8array"),
    await viewRoundTripZip.file(partPath).async("uint8array"),
    `only ppt/viewProps.xml may change (${partPath})`,
  );
}
const editedViewXml = await viewEditedZip.file("ppt/viewProps.xml").async("text");
assert.match(editedViewXml, /cx="72009"/);
assert.match(editedViewXml, /cy="91441"/);
assert.match(editedViewXml, /snapToGrid="0"/);
assert.match(editedViewXml, /snapToObjects="1"/);
assert.match(editedViewXml, /showGuides="1"/);
assert.match(editedViewXml, /orient="horz" pos="2161"/);
assert.match(editedViewXml, /orient="vert" pos="2881"/);
const editedViewRoundTrip = await PresentationFile.importPptx(viewEditedFile);
assert.equal(editedViewRoundTrip.view.gridSpacingCxEmu, 72_009);
assert.equal(editedViewRoundTrip.view.gridSpacingCyEmu, 91_441);
assert.equal(editedViewRoundTrip.view.slideViewSnapToGrid, false);
assert.equal(editedViewRoundTrip.view.slideViewSnapToObjects, true);
assert.deepEqual(editedViewRoundTrip.view.slideGuides, [
  { orientation: "horizontal", position: 2161 },
  { orientation: "vertical", position: 2881 },
]);
const tamperedViewState = editedViewRoundTrip[Symbol.for("office-kit.presentation-state")];
tamperedViewState.viewProperties.source.residualSha256 = "0".repeat(64);
await assert.rejects(
  () => PresentationFile.exportPptx(editedViewRoundTrip),
  (error) => error?.code === "presentation_view_source_binding_mismatch",
);

// Eligible imported top-level OLE objects expose one deliberately narrow edit:
// replacing the uniquely bound XLSX payload. The OLE shell, preview image,
// relationships, source package, and every other native part stay source-owned.
const embeddedSourceWorkbook = Workbook.create();
embeddedSourceWorkbook.worksheets.add("Embedded").getRange("A1").values = [["Original embedded workbook"]];
const embeddedSourceXlsx = await SpreadsheetFile.exportXlsx(embeddedSourceWorkbook);
const embeddedPreviewBytes = Buffer.from(PNG.split(",")[1], "base64");
const oleFrame = '<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="100" name="Embedded workbook"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2286000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj showAsIcon="1" r:id="rIdEmbeddedWorkbook" imgW="965200" imgH="609600" progId="Excel.Sheet.12"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdEmbeddedPreview"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2286000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>';
const firstSlideRelationships = await firstZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const oleSource = await PresentationFile.patchPptx(firstExport, [
  { path: "ppt/slides/slide1.xml", xml: firstSlideXml.replace("</p:spTree>", `${oleFrame}</p:spTree>`) },
  { path: "ppt/slides/_rels/slide1.xml.rels", xml: firstSlideRelationships.replace("</Relationships>", '<Relationship Id="rIdEmbeddedWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/agent-workbook.xlsx"/><Relationship Id="rIdEmbeddedPreview" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/agent-workbook-preview.png"/></Relationships>') },
  { path: "ppt/embeddings/agent-workbook.xlsx", bytes: embeddedSourceXlsx.bytes, contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  { path: "ppt/media/agent-workbook-preview.png", bytes: embeddedPreviewBytes, contentType: "image/png" },
]);
const oleSourceSnapshot = Uint8Array.from(oleSource.bytes);
const olePresentation = await PresentationFile.importPptx(oleSource);
const oleObject = itemByName(olePresentation.slides.getItem(0).nativeObjects.items, "Embedded workbook");
assert.equal(oleObject.nativeKind, "oleObject");
assert.equal(oleObject.editable, false);
assert.deepEqual(oleObject.inspectRecord().editableFields, ["embeddedWorkbook"]);
assert.throws(() => { oleObject.oleWorkbook = undefined; }, TypeError);
assert.throws(() => oleObject.setName("Unsafe shell rename"), /read-only/);
assert.throws(() => oleObject.replaceEmbeddedWorkbook("not bytes"), /FileBlob, Uint8Array, ArrayBuffer/);
assert.throws(() => oleObject.replaceEmbeddedWorkbook(new Uint8Array()), /1 through 16777216 bytes/);
assert.throws(() => oleObject.replaceEmbeddedWorkbook(new Uint8Array(16 * 1024 * 1024 + 1)), /1 through 16777216 bytes/);
const extractedSourceWorkbook = await SpreadsheetFile.importXlsx(oleObject.getEmbeddedWorkbook());
assert.equal(extractedSourceWorkbook.worksheets.getItem("Embedded").getRange("A1").values[0][0], "Original embedded workbook");
const extractedGenericWorkbook = oleObject.getEmbeddedOfficePackage();
assert.equal(extractedGenericWorkbook.metadata.artifactKind, "officePackage");
assert.equal(extractedGenericWorkbook.metadata.officePackageKind, "xlsx");
assert.equal((await SpreadsheetFile.importXlsx(extractedGenericWorkbook)).worksheets.getItem("Embedded").getRange("A1").values[0][0], "Original embedded workbook");

const embeddedReplacementWorkbook = Workbook.create();
embeddedReplacementWorkbook.worksheets.add("Embedded").getRange("A1:B2").values = [["Replacement workbook", 42], ["Verified", true]];
const embeddedReplacementXlsx = await SpreadsheetFile.exportXlsx(embeddedReplacementWorkbook);
const mutableReplacement = Uint8Array.from(embeddedReplacementXlsx.bytes);
assert.equal(oleObject.replaceEmbeddedWorkbook(mutableReplacement), oleObject);
mutableReplacement.fill(0);
const pendingWorkbookFile = oleObject.getEmbeddedWorkbook();
assert.equal(pendingWorkbookFile.metadata.pendingReplacement, true);
pendingWorkbookFile.bytes.fill(0);
const pendingWorkbook = await SpreadsheetFile.importXlsx(oleObject.getEmbeddedWorkbook());
assert.deepEqual(pendingWorkbook.worksheets.getItem("Embedded").getRange("A1:B2").values, [["Replacement workbook", 42], ["Verified", true]]);
const replacementView = new DataView(embeddedReplacementXlsx.bytes.buffer, embeddedReplacementXlsx.bytes.byteOffset, embeddedReplacementXlsx.bytes.byteLength);
assert.equal(oleObject.replaceEmbeddedWorkbook(replacementView), oleObject);
assert.equal(oleObject.replaceEmbeddedOfficePackage(embeddedReplacementXlsx), oleObject);
assert.match(olePresentation.inspect({ kind: "nativeObject", target: oleObject.id, maxChars: 4000 }).ndjson, /"replacementPending":true/);

const oleExport = await PresentationFile.exportPptx(olePresentation);
assert.deepEqual(oleSource.bytes, oleSourceSnapshot);
const oleSourceZipForComparison = await JSZip.loadAsync(oleSource.bytes);
const oleOutputZip = await JSZip.loadAsync(oleExport.bytes);
assert.deepEqual(Object.keys(oleOutputZip.files).sort(), Object.keys(oleSourceZipForComparison.files).sort());
for (const partPath of Object.keys(oleSourceZipForComparison.files)) {
  if (oleSourceZipForComparison.files[partPath].dir || partPath === "ppt/embeddings/agent-workbook.xlsx") continue;
  assert.deepEqual(
    await oleOutputZip.file(partPath).async("uint8array"),
    await oleSourceZipForComparison.file(partPath).async("uint8array"),
    `OLE payload replacement must preserve ${partPath} byte-for-byte`,
  );
}
assert.deepEqual(await oleOutputZip.file("ppt/embeddings/agent-workbook.xlsx").async("uint8array"), embeddedReplacementXlsx.bytes);
assert.deepEqual(await oleOutputZip.file("ppt/media/agent-workbook-preview.png").async("uint8array"), Uint8Array.from(embeddedPreviewBytes));
assert.match(await oleOutputZip.file("ppt/slides/slide1.xml").async("text"), /r:id="rIdEmbeddedWorkbook"/);
const oleRoundTrip = await PresentationFile.importPptx(oleExport);
const reboundOleObject = itemByName(oleRoundTrip.slides.getItem(0).nativeObjects.items, "Embedded workbook");
assert.equal(reboundOleObject.inspectRecord().embeddedWorkbook.replacementPending, false);
assert.notEqual(reboundOleObject.oleWorkbook.sourceSha256, oleObject.oleWorkbook.sourceSha256);
const reboundWorkbook = await SpreadsheetFile.importXlsx(reboundOleObject.getEmbeddedWorkbook());
assert.deepEqual(reboundWorkbook.worksheets.getItem("Embedded").getRange("A1:B2").values, [["Replacement workbook", 42], ["Verified", true]]);

const invalidOlePresentation = await PresentationFile.importPptx(oleSource);
const invalidOleObject = itemByName(invalidOlePresentation.slides.getItem(0).nativeObjects.items, "Embedded workbook");
invalidOleObject.replaceEmbeddedWorkbook(Uint8Array.from([0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4]));
await assert.rejects(
  () => PresentationFile.exportPptx(invalidOlePresentation),
  (error) => new Set(["invalid_opc_package", "invalid_presentation_ole_workbook"]).has(error?.code),
);

const oleSourceZip = await JSZip.loadAsync(oleSource.bytes);
const oleSourceRelationships = await oleSourceZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const sharedOleSource = await PresentationFile.patchPptx(oleSource, [{
  path: "ppt/slides/_rels/slide1.xml.rels",
  xml: oleSourceRelationships.replace("</Relationships>", '<Relationship Id="rIdSharedEmbeddedWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/agent-workbook.xlsx"/></Relationships>'),
}]);
const sharedOlePresentation = await PresentationFile.importPptx(sharedOleSource);
const sharedOleObject = itemByName(sharedOlePresentation.slides.getItem(0).nativeObjects.items, "Embedded workbook");
assert.equal(sharedOleObject.oleWorkbook, undefined);
assert.deepEqual(sharedOleObject.inspectRecord().editableFields, []);
assert.throws(() => sharedOleObject.getEmbeddedWorkbook(), /has no embedded XLSX workbook/);
assert.equal(sharedOlePresentation.slides.getItem(0).cloneCapability.supported, true);
const sharedOleSlideCount = sharedOlePresentation.slides.items.length;
sharedOlePresentation.slides.getItem(0).duplicate();
assert.equal((await PresentationFile.importPptx(await PresentationFile.exportPptx(sharedOlePresentation))).slides.items.length, sharedOleSlideCount + 1);

// The additive Office-package capability is not a generic OLE escape hatch:
// today it recognizes only one uniquely-bound DOCX payload. It preserves the
// existing OLE shell/preview/relationship graph, and DOCX OLE frames cannot
// enter the XLSX-only slide-clone profile.
const embeddedSourceDocument = DocumentModel.create({ name: "Original embedded document", blocks: [] });
embeddedSourceDocument.addParagraph("Original embedded document marker");
const embeddedSourceDocx = await DocumentFile.exportDocx(embeddedSourceDocument);
const docxContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const oleDocumentFrame = '<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="101" name="Embedded document"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="1828800"/><a:ext cx="3657600" cy="2286000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj showAsIcon="1" r:id="rIdEmbeddedDocument" imgW="965200" imgH="609600" progId="Word.Document.12"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdEmbeddedDocumentPreview"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="3657600" cy="2286000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>';
const oleDocumentSource = await PresentationFile.patchPptx(firstExport, [
  { path: "ppt/slides/slide1.xml", xml: firstSlideXml.replace("</p:spTree>", `${oleDocumentFrame}</p:spTree>`) },
  { path: "ppt/slides/_rels/slide1.xml.rels", xml: firstSlideRelationships.replace("</Relationships>", '<Relationship Id="rIdEmbeddedDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/agent-document.docx"/><Relationship Id="rIdEmbeddedDocumentPreview" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/agent-document-preview.png"/></Relationships>') },
  { path: "ppt/embeddings/agent-document.docx", bytes: embeddedSourceDocx.bytes, contentType: docxContentType },
  { path: "ppt/media/agent-document-preview.png", bytes: embeddedPreviewBytes, contentType: "image/png" },
]);
const oleDocumentSourceSnapshot = Uint8Array.from(oleDocumentSource.bytes);
const oleDocumentPresentation = await PresentationFile.importPptx(oleDocumentSource);
const oleDocumentObject = itemByName(oleDocumentPresentation.slides.getItem(0).nativeObjects.items, "Embedded document");
assert.equal(oleDocumentObject.nativeKind, "oleObject");
assert.equal(oleDocumentObject.oleWorkbook, undefined);
assert.equal(oleDocumentObject.oleOfficePackage.kind, "docx");
assert.equal(oleDocumentObject.oleOfficePackage.contentType, docxContentType);
assert.deepEqual(oleDocumentObject.inspectRecord().editableFields, ["embeddedOfficePackage"]);
assert.equal(oleDocumentObject.inspectRecord().embeddedOfficePackage.kind, "docx");
const extractedSourceDocument = await DocumentFile.importDocx(oleDocumentObject.getEmbeddedOfficePackage());
assert.deepEqual(extractedSourceDocument.paragraphs, ["Original embedded document marker"]);
assert.throws(() => oleDocumentObject.replaceEmbeddedOfficePackage("not bytes"), /FileBlob, Uint8Array, ArrayBuffer/);
assert.throws(() => oleDocumentObject.replaceEmbeddedOfficePackage(new FileBlob(embeddedSourceDocx.bytes, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })), /retain content type/);
assert.throws(() => oleDocumentObject.replaceEmbeddedOfficePackage(new Uint8Array()), /1 through 16777216 bytes/);
assert.throws(() => oleDocumentObject.replaceEmbeddedOfficePackage(new Uint8Array(16 * 1024 * 1024 + 1)), /1 through 16777216 bytes/);

const embeddedReplacementDocument = DocumentModel.create({ name: "Replacement embedded document", blocks: [] });
embeddedReplacementDocument.addParagraph("Replacement embedded document marker");
const embeddedReplacementDocx = await DocumentFile.exportDocx(embeddedReplacementDocument);
const mutableDocumentReplacement = Uint8Array.from(embeddedReplacementDocx.bytes);
assert.equal(oleDocumentObject.replaceEmbeddedOfficePackage(new FileBlob(mutableDocumentReplacement, { type: docxContentType })), oleDocumentObject);
mutableDocumentReplacement.fill(0);
const pendingDocumentFile = oleDocumentObject.getEmbeddedOfficePackage();
assert.equal(pendingDocumentFile.metadata.pendingReplacement, true);
pendingDocumentFile.bytes.fill(0);
assert.deepEqual((await DocumentFile.importDocx(oleDocumentObject.getEmbeddedOfficePackage())).paragraphs, ["Replacement embedded document marker"]);
assert.match(oleDocumentPresentation.inspect({ kind: "nativeObject", target: oleDocumentObject.id, maxChars: 4000 }).ndjson, /"embeddedOfficePackage"/);

const oleDocumentExport = await PresentationFile.exportPptx(oleDocumentPresentation);
assert.deepEqual(oleDocumentSource.bytes, oleDocumentSourceSnapshot);
const oleDocumentSourceZip = await JSZip.loadAsync(oleDocumentSource.bytes);
const oleDocumentOutputZip = await JSZip.loadAsync(oleDocumentExport.bytes);
assert.deepEqual(Object.keys(oleDocumentOutputZip.files).sort(), Object.keys(oleDocumentSourceZip.files).sort());
for (const partPath of Object.keys(oleDocumentSourceZip.files)) {
  if (oleDocumentSourceZip.files[partPath].dir || partPath === "ppt/embeddings/agent-document.docx") continue;
  assert.deepEqual(
    await oleDocumentOutputZip.file(partPath).async("uint8array"),
    await oleDocumentSourceZip.file(partPath).async("uint8array"),
    `DOCX OLE payload replacement must preserve ${partPath} byte-for-byte`,
  );
}
assert.deepEqual(await oleDocumentOutputZip.file("ppt/embeddings/agent-document.docx").async("uint8array"), embeddedReplacementDocx.bytes);
assert.deepEqual(await oleDocumentOutputZip.file("ppt/media/agent-document-preview.png").async("uint8array"), Uint8Array.from(embeddedPreviewBytes));
assert.match(await oleDocumentOutputZip.file("ppt/slides/slide1.xml").async("text"), /r:id="rIdEmbeddedDocument"/);
const oleDocumentRoundTrip = await PresentationFile.importPptx(oleDocumentExport);
const reboundDocumentObject = itemByName(oleDocumentRoundTrip.slides.getItem(0).nativeObjects.items, "Embedded document");
assert.equal(reboundDocumentObject.inspectRecord().embeddedOfficePackage.replacementPending, false);
assert.notEqual(reboundDocumentObject.oleOfficePackage.sourceSha256, oleDocumentObject.oleOfficePackage.sourceSha256);
assert.deepEqual((await DocumentFile.importDocx(reboundDocumentObject.getEmbeddedOfficePackage())).paragraphs, ["Replacement embedded document marker"]);

const invalidDocumentOlePresentation = await PresentationFile.importPptx(oleDocumentSource);
const invalidDocumentOleObject = itemByName(invalidDocumentOlePresentation.slides.getItem(0).nativeObjects.items, "Embedded document");
invalidDocumentOleObject.replaceEmbeddedOfficePackage(embeddedReplacementXlsx.bytes);
await assert.rejects(
  () => PresentationFile.exportPptx(invalidDocumentOlePresentation),
  (error) => new Set(["invalid_opc_package", "invalid_presentation_ole_office_package"]).has(error?.code),
);

const oleDocumentSourceZipForSharing = await JSZip.loadAsync(oleDocumentSource.bytes);
const sharedOleDocumentSource = await PresentationFile.patchPptx(oleDocumentSource, [{
  path: "ppt/slides/_rels/slide1.xml.rels",
  xml: (await oleDocumentSourceZipForSharing.file("ppt/slides/_rels/slide1.xml.rels").async("text")).replace("</Relationships>", '<Relationship Id="rIdSharedEmbeddedDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/agent-document.docx"/></Relationships>'),
}]);
const sharedOleDocumentPresentation = await PresentationFile.importPptx(sharedOleDocumentSource);
const sharedOleDocumentObject = itemByName(sharedOleDocumentPresentation.slides.getItem(0).nativeObjects.items, "Embedded document");
assert.equal(sharedOleDocumentObject.oleOfficePackage, undefined);
assert.deepEqual(sharedOleDocumentObject.inspectRecord().editableFields, []);
assert.throws(() => sharedOleDocumentObject.getEmbeddedOfficePackage(), /has no bounded embedded Office package/);

// The bounded imported-slide clone may carry the same uniquely bound,
// top-level embedded-XLSX OLE frame. The mutable workbook package is copied
// into a distinct part, while the immutable preview ImagePart is shared.
const oleCloneBaseSlideXml = await cloneSourceZip.file("ppt/slides/slide1.xml").async("text");
const oleCloneBaseRelationships = await cloneSourceZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const oleCloneSource = await PresentationFile.patchPptx(cloneSourcePptx, [
  { path: "ppt/slides/slide1.xml", xml: oleCloneBaseSlideXml.replace("</p:spTree>", `${oleFrame}</p:spTree>`) },
  { path: "ppt/slides/_rels/slide1.xml.rels", xml: oleCloneBaseRelationships.replace("</Relationships>", '<Relationship Id="rIdEmbeddedWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/clone-agent-workbook.xlsx"/><Relationship Id="rIdEmbeddedPreview" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/clone-agent-workbook-preview.png"/></Relationships>') },
  { path: "ppt/embeddings/clone-agent-workbook.xlsx", bytes: embeddedSourceXlsx.bytes, contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  { path: "ppt/media/clone-agent-workbook-preview.png", bytes: embeddedPreviewBytes, contentType: "image/png" },
]);
const oleCloneSourceSnapshot = Uint8Array.from(oleCloneSource.bytes);
const oleCloneImported = await PresentationFile.importPptx(oleCloneSource);
const oleCloneOrigin = oleCloneImported.slides.getItem(0);
const oleCloneOriginObject = itemByName(oleCloneOrigin.nativeObjects.items, "Embedded workbook");
const oleClonePending = oleCloneOrigin.duplicate();
const oleClonePendingObject = itemByName(oleClonePending.nativeObjects.items, "Embedded workbook");
assert.notEqual(oleClonePendingObject, oleCloneOriginObject);
assert.notEqual(oleClonePendingObject.id, oleCloneOriginObject.id);
assert.equal(oleClonePendingObject.oleWorkbook.partPath, oleCloneOriginObject.oleWorkbook.partPath);
assert.equal(oleClonePendingObject.oleWorkbook.sourceSha256, oleCloneOriginObject.oleWorkbook.sourceSha256);

const oleCloneExport = await PresentationFile.exportPptx(oleCloneImported);
assert.deepEqual(oleCloneSource.bytes, oleCloneSourceSnapshot);
const oleCloneSourceZip = await JSZip.loadAsync(oleCloneSource.bytes);
const oleCloneOutputZip = await JSZip.loadAsync(oleCloneExport.bytes);
const oleCloneSlidePath = (await orderedPptxSlidePaths(oleCloneOutputZip))[1];
assert.deepEqual(
  await oleCloneOutputZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await oleCloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "OLE cloning must retain the origin SlidePart byte-for-byte",
);
assert.deepEqual(
  await oleCloneOutputZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  await oleCloneSourceZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  "OLE cloning must retain the origin relationship part byte-for-byte",
);
const oleCloneWorkbookPaths = Object.keys(oleCloneOutputZip.files)
  .filter((partPath) => /^ppt\/(?:slides\/)?embeddings\/[^/]+\.xlsx$/i.test(partPath))
  .sort();
assert.equal(oleCloneWorkbookPaths.length, 2, "the clone must allocate exactly one additional XLSX package part");
const oleCloneWorkbookPart = oleCloneWorkbookPaths.find((partPath) => partPath !== "ppt/embeddings/clone-agent-workbook.xlsx");
assert.ok(oleCloneWorkbookPart);
assert.deepEqual(
  await oleCloneOutputZip.file(oleCloneWorkbookPart).async("uint8array"),
  await oleCloneSourceZip.file("ppt/embeddings/clone-agent-workbook.xlsx").async("uint8array"),
  "the first clone export must copy the embedded XLSX bytes exactly",
);
const relationshipForType = (xml, typeSuffix) => {
  const matches = [...xml.matchAll(/<Relationship\b[^>]*>/g)]
    .map(([tag]) => ({
      id: /\bId="([^"]+)"/.exec(tag)?.[1],
      type: /\bType="([^"]+)"/.exec(tag)?.[1],
      target: /\bTarget="([^"]+)"/.exec(tag)?.[1],
      targetMode: /\bTargetMode="([^"]+)"/.exec(tag)?.[1],
    }))
    .filter((relationship) => relationship.type?.endsWith(`/${typeSuffix}`));
  assert.equal(matches.length, 1, `expected one ${typeSuffix} relationship in ${xml}`);
  return matches[0];
};
const resolveSlideRelationshipTarget = (target) => target.startsWith("/")
  ? target.replace(/^\/+/, "")
  : path.posix.normalize(path.posix.join("ppt/slides", target));
const oleCloneSourceRelationships = await oleCloneOutputZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const oleCloneCopyRelationships = await oleCloneOutputZip.file(relationshipPartPath(oleCloneSlidePath)).async("text");
const oleCloneSourcePackageRelationship = relationshipForType(oleCloneSourceRelationships, "package");
const oleCloneCopyPackageRelationship = relationshipForType(oleCloneCopyRelationships, "package");
assert.equal(oleCloneCopyPackageRelationship.id, oleCloneSourcePackageRelationship.id);
assert.equal(oleCloneCopyPackageRelationship.targetMode, undefined);
assert.equal(resolveSlideRelationshipTarget(oleCloneSourcePackageRelationship.target), "ppt/embeddings/clone-agent-workbook.xlsx");
assert.equal(resolveSlideRelationshipTarget(oleCloneCopyPackageRelationship.target), oleCloneWorkbookPart);
const oleCloneSourcePreviewRelationship = relationshipForType(oleCloneSourceRelationships, "image");
const oleCloneCopyPreviewRelationship = relationshipForType(oleCloneCopyRelationships, "image");
assert.equal(oleCloneCopyPreviewRelationship.id, oleCloneSourcePreviewRelationship.id);
assert.equal(
  resolveSlideRelationshipTarget(oleCloneCopyPreviewRelationship.target),
  resolveSlideRelationshipTarget(oleCloneSourcePreviewRelationship.target),
  "both OLE frames must share the same immutable preview ImagePart",
);

const oleCloneRoundTrip = await PresentationFile.importPptx(oleCloneExport);
const oleCloneRoundTripOrigin = itemByName(oleCloneRoundTrip.slides.getItem(0).nativeObjects.items, "Embedded workbook");
const oleCloneRoundTripCopy = itemByName(oleCloneRoundTrip.slides.getItem(1).nativeObjects.items, "Embedded workbook");
assert.notEqual(oleCloneRoundTripCopy.oleWorkbook.partPath, oleCloneRoundTripOrigin.oleWorkbook.partPath);
assert.equal(oleCloneRoundTripCopy.oleWorkbook.sourceSha256, oleCloneRoundTripOrigin.oleWorkbook.sourceSha256);
oleCloneRoundTripCopy.replaceEmbeddedWorkbook(embeddedReplacementXlsx);
const oleCloneEditedExport = await PresentationFile.exportPptx(oleCloneRoundTrip);
const oleCloneEditedZip = await JSZip.loadAsync(oleCloneEditedExport.bytes);
assert.deepEqual(
  await oleCloneEditedZip.file(oleCloneRoundTripOrigin.oleWorkbook.partPath).async("uint8array"),
  embeddedSourceXlsx.bytes,
  "editing the reimported clone workbook must leave the origin package byte-for-byte intact",
);
assert.deepEqual(
  await oleCloneEditedZip.file(oleCloneRoundTripCopy.oleWorkbook.partPath).async("uint8array"),
  embeddedReplacementXlsx.bytes,
);
const oleCloneEditedRoundTrip = await PresentationFile.importPptx(oleCloneEditedExport);
const oleCloneEditedOriginWorkbook = await SpreadsheetFile.importXlsx(itemByName(oleCloneEditedRoundTrip.slides.getItem(0).nativeObjects.items, "Embedded workbook").getEmbeddedWorkbook());
const oleCloneEditedCopyWorkbook = await SpreadsheetFile.importXlsx(itemByName(oleCloneEditedRoundTrip.slides.getItem(1).nativeObjects.items, "Embedded workbook").getEmbeddedWorkbook());
assert.equal(oleCloneEditedOriginWorkbook.worksheets.getItem("Embedded").getRange("A1").values[0][0], "Original embedded workbook");
assert.equal(oleCloneEditedCopyWorkbook.worksheets.getItem("Embedded").getRange("A1").values[0][0], "Replacement workbook");

const immediateOleCloneEdit = await PresentationFile.importPptx(oleCloneSource);
itemByName(immediateOleCloneEdit.slides.getItem(0).duplicate().nativeObjects.items, "Embedded workbook")
  .replaceEmbeddedWorkbook(embeddedReplacementXlsx);
await assert.rejects(
  () => PresentationFile.exportPptx(immediateOleCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "a cloned OLE payload may be edited only after export and reimport establish independent source identity",
);

// A canonical top-level SmartArt frame owns exactly the four standard
// relationship-free DrawingML diagram roots. The bounded clone copies all
// four into distinct parts so a later source-bound edit cannot couple the
// origin and clone through shared diagram state.
const smartArtFrame = '<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="120" name="Clone-safe SmartArt"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="1828800"/><a:ext cx="5486400" cy="2743200"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdAgentDiagramData" r:lo="rIdAgentDiagramLayout" r:qs="rIdAgentDiagramStyle" r:cs="rIdAgentDiagramColors"/></a:graphicData></a:graphic></p:graphicFrame>';
const smartArtRelationships = '<Relationship Id="rIdAgentDiagramData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/agent-data.xml"/><Relationship Id="rIdAgentDiagramLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/agent-layout.xml"/><Relationship Id="rIdAgentDiagramStyle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/agent-style.xml"/><Relationship Id="rIdAgentDiagramColors" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/agent-colors.xml"/>';
const smartArtParts = [
  ["ppt/diagrams/agent-data.xml", "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml", '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:ptLst/><dgm:cxnLst/><dgm:bg/><dgm:whole/></dgm:dataModel>'],
  ["ppt/diagrams/agent-layout.xml", "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml", '<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:office-kit:agent-layout"><dgm:title val="Agent"/><dgm:desc val="Agent layout"/><dgm:catLst/><dgm:layoutNode name="root"/></dgm:layoutDef>'],
  ["ppt/diagrams/agent-style.xml", "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml", '<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:office-kit:agent-style"><dgm:title val="Agent"/><dgm:desc val="Agent style"/><dgm:catLst/><dgm:styleLbl name="node0"/></dgm:styleDef>'],
  ["ppt/diagrams/agent-colors.xml", "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml", '<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:office-kit:agent-colors"><dgm:title val="Agent"/><dgm:desc val="Agent colors"/><dgm:catLst/></dgm:colorsDef>'],
];
const smartArtSource = await PresentationFile.patchPptx(cloneSourcePptx, [
  { path: "ppt/slides/slide1.xml", xml: oleCloneBaseSlideXml.replace("</p:spTree>", `${smartArtFrame}</p:spTree>`) },
  { path: "ppt/slides/_rels/slide1.xml.rels", xml: oleCloneBaseRelationships.replace("</Relationships>", `${smartArtRelationships}</Relationships>`) },
  ...smartArtParts.map(([partPath, contentType, xml]) => ({ path: partPath, contentType, xml })),
]);
const smartArtSourceSnapshot = Uint8Array.from(smartArtSource.bytes);
const smartArtImported = await PresentationFile.importPptx(smartArtSource);
const smartArtOriginSlide = smartArtImported.slides.getItem(0);
const smartArtOrigin = itemByName(smartArtOriginSlide.nativeObjects.items, "Clone-safe SmartArt");
assert.equal(smartArtOrigin.nativeKind, "diagram");
assert.equal(smartArtOrigin.parts.length, 4);
assert.ok(smartArtOrigin.parts.every((part) => part.relationships.length === 0));
const smartArtPendingSlide = smartArtOriginSlide.duplicate();
const smartArtPending = itemByName(smartArtPendingSlide.nativeObjects.items, "Clone-safe SmartArt");
assert.notEqual(smartArtPending, smartArtOrigin);
assert.notEqual(smartArtPending.id, smartArtOrigin.id);
assert.deepEqual(smartArtPending.parts.map((part) => part.path), smartArtOrigin.parts.map((part) => part.path));

const smartArtExport = await PresentationFile.exportPptx(smartArtImported);
assert.deepEqual(smartArtSource.bytes, smartArtSourceSnapshot);
const smartArtSourceZip = await JSZip.loadAsync(smartArtSource.bytes);
const smartArtOutputZip = await JSZip.loadAsync(smartArtExport.bytes);
const smartArtCloneSlidePath = (await orderedPptxSlidePaths(smartArtOutputZip))[1];
assert.deepEqual(
  await smartArtOutputZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await smartArtSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "SmartArt cloning must retain the origin SlidePart byte-for-byte",
);
assert.deepEqual(
  await smartArtOutputZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  await smartArtSourceZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  "SmartArt cloning must retain the origin relationship part byte-for-byte",
);
const smartArtSourceRels = await smartArtOutputZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const smartArtCloneRels = await smartArtOutputZip.file(relationshipPartPath(smartArtCloneSlidePath)).async("text");
const smartArtTypeSuffixes = ["diagramData", "diagramLayout", "diagramQuickStyle", "diagramColors"];
const smartArtClonePartPaths = [];
for (const typeSuffix of smartArtTypeSuffixes) {
  const sourceRelationship = relationshipForType(smartArtSourceRels, typeSuffix);
  const cloneRelationship = relationshipForType(smartArtCloneRels, typeSuffix);
  const sourcePath = resolveSlideRelationshipTarget(sourceRelationship.target);
  const clonePath = resolveSlideRelationshipTarget(cloneRelationship.target);
  assert.equal(cloneRelationship.id, sourceRelationship.id);
  assert.notEqual(clonePath, sourcePath);
  assert.deepEqual(
    await smartArtOutputZip.file(clonePath).async("uint8array"),
    await smartArtSourceZip.file(sourcePath).async("uint8array"),
    `SmartArt cloning must byte-copy the closed ${typeSuffix} part`,
  );
  smartArtClonePartPaths.push(clonePath);
}
assert.equal(new Set(smartArtClonePartPaths).size, 4);

const smartArtRoundTrip = await PresentationFile.importPptx(smartArtExport);
const smartArtRoundTripOrigin = itemByName(smartArtRoundTrip.slides.getItem(0).nativeObjects.items, "Clone-safe SmartArt");
const smartArtRoundTripClone = itemByName(smartArtRoundTrip.slides.getItem(1).nativeObjects.items, "Clone-safe SmartArt");
assert.equal(smartArtRoundTripOrigin.parts.length, 4);
assert.equal(smartArtRoundTripClone.parts.length, 4);
assert.equal(
  smartArtRoundTripOrigin.parts.some((part) => smartArtRoundTripClone.parts.some((clonePart) => clonePart.path === part.path)),
  false,
  "reimported SmartArt origin and clone must not share any mutable diagram part",
);
assert.deepEqual(
  smartArtRoundTripOrigin.parts.map((part) => part.sourceSha256).sort(),
  smartArtRoundTripClone.parts.map((part) => part.sourceSha256).sort(),
);

// A canonical closed SmartArt data model can expose only its direct plain
// document-node text. The bounded edit rewrites the one hash-bound data part;
// the frame, relationships, and layout/style/color leaves must stay intact.
const smartArtTextData = '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="{B31B1833-2B65-4D6B-B3D4-9B3988427B21}" type="doc"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Original node</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="1" type="doc"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Second node</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst/><dgm:bg/><dgm:whole/></dgm:dataModel>';
const smartArtTextSource = await PresentationFile.patchPptx(cloneSourcePptx, [
  { path: "ppt/slides/slide1.xml", xml: oleCloneBaseSlideXml.replace("</p:spTree>", `${smartArtFrame}</p:spTree>`) },
  { path: "ppt/slides/_rels/slide1.xml.rels", xml: oleCloneBaseRelationships.replace("</Relationships>", `${smartArtRelationships}</Relationships>`) },
  ...smartArtParts.map(([partPath, contentType, xml]) => ({ path: partPath, contentType, xml: partPath === "ppt/diagrams/agent-data.xml" ? smartArtTextData : xml })),
]);
const smartArtTextInput = Uint8Array.from(smartArtTextSource.bytes);
const smartArtTextImported = await PresentationFile.importPptx(smartArtTextSource);
const smartArtTextObject = itemByName(smartArtTextImported.slides.getItem(0).nativeObjects.items, "Clone-safe SmartArt");
assert.equal(smartArtTextObject.editable, false);
assert.deepEqual(smartArtTextObject.diagramText?.nodes, [
  { id: "{B31B1833-2B65-4D6B-B3D4-9B3988427B21}", text: "Original node", runs: ["Original node"] },
  { id: "1", text: "Second node", runs: ["Second node"] },
]);
assert.deepEqual(smartArtTextObject.inspectRecord().editableFields, ["diagramText"]);
assert.throws(
  () => smartArtTextObject.setDiagramNodeText("missing", "nope"),
  /not part of the source-bound diagram profile/,
);
assert.throws(
  () => smartArtTextObject.setDiagramNodeText("{B31B1833-2B65-4D6B-B3D4-9B3988427B21}", "x".repeat(32_768)),
  /32767 XML-safe characters/,
);
smartArtTextObject.setDiagramNodeText("{B31B1833-2B65-4D6B-B3D4-9B3988427B21}", " Revised node ");
const smartArtTextSlideCount = smartArtTextImported.slides.items.length;
assert.throws(
  () => smartArtTextImported.slides.getItem(0).duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "a pending SmartArt text edit must cross an export/reimport boundary before cloning",
);
assert.equal(smartArtTextImported.slides.items.length, smartArtTextSlideCount);
const smartArtTextExport = await PresentationFile.exportPptx(smartArtTextImported);
assert.deepEqual(smartArtTextSource.bytes, smartArtTextInput, "SmartArt text edits must preserve the caller input bytes");
const smartArtTextSourceZip = await JSZip.loadAsync(smartArtTextSource.bytes);
const smartArtTextOutputZip = await JSZip.loadAsync(smartArtTextExport.bytes);
for (const path of [
  "ppt/slides/slide1.xml",
  "ppt/slides/_rels/slide1.xml.rels",
  "ppt/diagrams/agent-layout.xml",
  "ppt/diagrams/agent-style.xml",
  "ppt/diagrams/agent-colors.xml",
]) {
  assert.deepEqual(
    await smartArtTextOutputZip.file(path).async("uint8array"),
    await smartArtTextSourceZip.file(path).async("uint8array"),
    `SmartArt text edits must not alter ${path}`,
  );
}
const smartArtTextOutputData = await smartArtTextOutputZip.file("ppt/diagrams/agent-data.xml").async("text");
assert.match(smartArtTextOutputData, / Revised node /);
assert.match(smartArtTextOutputData, /xml:space="preserve"/);
assert.notDeepEqual(
  await smartArtTextOutputZip.file("ppt/diagrams/agent-data.xml").async("uint8array"),
  await smartArtTextSourceZip.file("ppt/diagrams/agent-data.xml").async("uint8array"),
);
const smartArtTextRoundTrip = await PresentationFile.importPptx(smartArtTextExport);
const smartArtTextRebound = itemByName(smartArtTextRoundTrip.slides.getItem(0).nativeObjects.items, "Clone-safe SmartArt");
assert.deepEqual(smartArtTextRebound.diagramText?.nodes, [
  { id: "{B31B1833-2B65-4D6B-B3D4-9B3988427B21}", text: " Revised node ", runs: [" Revised node "] },
  { id: "1", text: "Second node", runs: ["Second node"] },
]);
assert.notEqual(smartArtTextRebound.diagramText?.sourceSha256, smartArtTextObject.diagramText?.sourceSha256);

const richSmartArtTextSource = await PresentationFile.patchPptx(smartArtTextSource, [{
  path: "ppt/diagrams/agent-data.xml",
  xml: smartArtTextData.replace(
    "<a:r><a:t>Original node</a:t></a:r>",
    '<a:pPr algn="ctr"/><a:r><a:rPr b="1"/><a:t>Original</a:t></a:r><a:br><a:rPr lang="fr-FR"/></a:br><a:endParaRPr lang="en-US"/></a:p><a:p><a:pPr marL="91440"/><a:endParaRPr lang="de-DE"/></a:p><a:p><a:r><a:rPr i="1"/><a:t> node</a:t></a:r>',
  ),
}]);
const richSmartArtText = await PresentationFile.importPptx(richSmartArtTextSource);
const richSmartArtTextObject = itemByName(richSmartArtText.slides.getItem(0).nativeObjects.items, "Clone-safe SmartArt");
assert.deepEqual(richSmartArtTextObject.diagramText.nodes[0], {
  id: "{B31B1833-2B65-4D6B-B3D4-9B3988427B21}",
  text: "Original node",
  runs: ["Original", " node"],
});
assert.throws(
  () => richSmartArtTextObject.setDiagramNodeText("{B31B1833-2B65-4D6B-B3D4-9B3988427B21}", "Revised node"),
  /2 source-bound styled runs.*setDiagramNodeRunText/,
);
assert.throws(
  () => richSmartArtTextObject.setDiagramNodeRunText("{B31B1833-2B65-4D6B-B3D4-9B3988427B21}", 2, "missing"),
  /no source-bound run at index 2/,
);
richSmartArtTextObject.setDiagramNodeRunText("{B31B1833-2B65-4D6B-B3D4-9B3988427B21}", 0, "Revised");
const richSmartArtTextExport = await PresentationFile.exportPptx(richSmartArtText);
const richSmartArtTextOutputZip = await JSZip.loadAsync(richSmartArtTextExport.bytes);
const richSmartArtTextOutputData = await richSmartArtTextOutputZip.file("ppt/diagrams/agent-data.xml").async("text");
assert.equal((richSmartArtTextOutputData.match(/<a:p(?:\s|>)/g) || []).length, 4,
  "the edited three-paragraph node and untouched second node must retain all source paragraphs");
assert.match(richSmartArtTextOutputData, /<a:pPr algn="ctr"\s*\/>/);
assert.match(richSmartArtTextOutputData, /<a:p><a:pPr marL="91440"\s*\/><a:endParaRPr lang="de-DE"\s*\/><\/a:p>/);
assert.match(richSmartArtTextOutputData, /<a:br><a:rPr lang="fr-FR"\s*\/><\/a:br>/);
assert.match(richSmartArtTextOutputData, /<a:endParaRPr lang="en-US"\s*\/>/);
assert.match(richSmartArtTextOutputData, /<a:rPr b="1"\s*\/>/);
assert.match(richSmartArtTextOutputData, /<a:rPr i="1"\s*\/>/);
assert.match(richSmartArtTextOutputData, /<a:t>Revised<\/a:t>/);
assert.deepEqual(
  itemByName((await PresentationFile.importPptx(richSmartArtTextExport)).slides.getItem(0).nativeObjects.items, "Clone-safe SmartArt").diagramText.nodes[0],
  { id: "{B31B1833-2B65-4D6B-B3D4-9B3988427B21}", text: "Revised node", runs: ["Revised", " node"] },
);

const attributedSmartArtBreakSource = await PresentationFile.patchPptx(richSmartArtTextSource, [{
  path: "ppt/diagrams/agent-data.xml",
  xml: (await (await JSZip.loadAsync(richSmartArtTextSource.bytes)).file("ppt/diagrams/agent-data.xml").async("text"))
    .replace("<a:br><a:rPr", '<a:br dirty="1"><a:rPr'),
}]);
const attributedSmartArtBreak = await PresentationFile.importPptx(attributedSmartArtBreakSource);
assert.equal(
  itemByName(attributedSmartArtBreak.slides.getItem(0).nativeObjects.items, "Clone-safe SmartArt").diagramText,
  undefined,
  "a break with unsupported attributes must withhold the SmartArt text capability",
);

const invalidSmartArtModelIdSource = await PresentationFile.patchPptx(smartArtTextSource, [{
  path: "ppt/diagrams/agent-data.xml",
  xml: smartArtTextData.replace("{B31B1833-2B65-4D6B-B3D4-9B3988427B21}", "agent-node-1"),
}]);
const invalidSmartArtModelId = await PresentationFile.importPptx(invalidSmartArtModelIdSource);
assert.equal(itemByName(invalidSmartArtModelId.slides.getItem(0).nativeObjects.items, "Clone-safe SmartArt").diagramText, undefined,
  "an invalid ST_ModelId must not expose a SmartArt text-edit capability");

const connectedSmartArtSource = await PresentationFile.patchPptx(smartArtSource, [{
  path: "ppt/diagrams/_rels/agent-data.xml.rels",
  xml: '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdUnsafeSmartArtLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/smartart" TargetMode="External"/></Relationships>',
}]);
const connectedSmartArt = await PresentationFile.importPptx(connectedSmartArtSource);
assert.equal(connectedSmartArt.slides.getItem(0).cloneCapability.supported, true);
const connectedSmartArtSlideCount = connectedSmartArt.slides.items.length;
connectedSmartArt.slides.getItem(0).duplicate();
assert.equal((await PresentationFile.importPptx(await PresentationFile.exportPptx(connectedSmartArt))).slides.items.length, connectedSmartArtSlideCount + 1);

const nestedSmartArt = await PresentationFile.importPptx(smartArtSource);
const nestedSmartArtObject = itemByName(nestedSmartArt.slides.getItem(0).nativeObjects.items, "Clone-safe SmartArt");
nestedSmartArtObject.rawXml = nestedSmartArtObject.rawXml.replace(/^<p:graphicFrame/, "<p:grpSp");
const nestedSmartArtSlideCount = nestedSmartArt.slides.items.length;
assert.throws(
  () => nestedSmartArt.slides.getItem(0).duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "a SmartArt graph whose source binding is not a top-level graphicFrame must fail before mutating the model",
);
assert.equal(nestedSmartArt.slides.items.length, nestedSmartArtSlideCount);

// A canonical top-level p:contentPart is the PresentationML carrier for one
// standard InkML CustomXmlPart. The clone must allocate a new InkML part under
// the same slide-local r:id rather than sharing mutable ink XML.
const inkContentElement = '<p:contentPart xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdAgentInk"><p14:nvContentPartPr><p14:cNvPr id="121" name="Clone-safe ink"/><p14:cNvContentPartPr/><p14:nvPr/></p14:nvContentPartPr><p14:xfrm><a:off x="914400" y="1828800"/><a:ext cx="4572000" cy="2286000"/></p14:xfrm></p:contentPart>';
const inkRelationship = '<Relationship Id="rIdAgentInk" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/agent-ink.xml"/>';
const inkXml = '<ink xmlns="http://www.w3.org/2003/InkML"><trace>0 0, 100 100, 200 0</trace></ink>';
const inkSource = await PresentationFile.patchPptx(cloneSourcePptx, [
  { path: "ppt/slides/slide1.xml", xml: oleCloneBaseSlideXml.replace("</p:spTree>", `${inkContentElement}</p:spTree>`) },
  { path: "ppt/slides/_rels/slide1.xml.rels", xml: oleCloneBaseRelationships.replace("</Relationships>", `${inkRelationship}</Relationships>`) },
  { path: "ppt/customXml/agent-ink.xml", contentType: "application/inkml+xml", xml: inkXml },
]);
const inkSourceSnapshot = Uint8Array.from(inkSource.bytes);
const inkImported = await PresentationFile.importPptx(inkSource);
const inkOriginSlide = inkImported.slides.getItem(0);
const inkOrigin = itemByName(inkOriginSlide.nativeObjects.items, "Clone-safe ink");
assert.equal(inkOrigin.nativeKind, "contentPart");
assert.deepEqual(inkOrigin.position, { left: 96, top: 192, width: 480, height: 240 });
assert.equal(inkOrigin.parts.length, 1);
assert.equal(inkOrigin.parts[0].contentType, "application/inkml+xml");
assert.equal(inkOrigin.parts[0].relationships.length, 0);
const inkPendingSlide = inkOriginSlide.duplicate();
const inkPending = itemByName(inkPendingSlide.nativeObjects.items, "Clone-safe ink");
assert.notEqual(inkPending, inkOrigin);
assert.notEqual(inkPending.id, inkOrigin.id);
assert.equal(inkPending.parts[0].path, inkOrigin.parts[0].path);

const inkExport = await PresentationFile.exportPptx(inkImported);
assert.deepEqual(inkSource.bytes, inkSourceSnapshot);
const inkSourceZip = await JSZip.loadAsync(inkSource.bytes);
const inkOutputZip = await JSZip.loadAsync(inkExport.bytes);
const inkCloneSlidePath = (await orderedPptxSlidePaths(inkOutputZip))[1];
assert.deepEqual(
  await inkOutputZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await inkSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "InkML cloning must retain the origin SlidePart byte-for-byte",
);
assert.deepEqual(
  await inkOutputZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  await inkSourceZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  "InkML cloning must retain the origin relationship part byte-for-byte",
);
const inkSourceRelationship = relationshipForType(await inkOutputZip.file("ppt/slides/_rels/slide1.xml.rels").async("text"), "customXml");
const inkCloneRelationship = relationshipForType(await inkOutputZip.file(relationshipPartPath(inkCloneSlidePath)).async("text"), "customXml");
const inkSourcePartPath = resolveSlideRelationshipTarget(inkSourceRelationship.target);
const inkClonePartPath = resolveSlideRelationshipTarget(inkCloneRelationship.target);
assert.equal(inkCloneRelationship.id, inkSourceRelationship.id);
assert.equal(inkSourcePartPath, "ppt/customXml/agent-ink.xml");
assert.match(inkClonePartPath, /^ppt\/customXml\/[^/]+\.xml$/i);
assert.notEqual(inkClonePartPath, inkSourcePartPath);
assert.deepEqual(
  await inkOutputZip.file(inkClonePartPath).async("uint8array"),
  await inkSourceZip.file(inkSourcePartPath).async("uint8array"),
  "InkML cloning must byte-copy the closed content part",
);

const inkRoundTrip = await PresentationFile.importPptx(inkExport);
const inkRoundTripOrigin = itemByName(inkRoundTrip.slides.getItem(0).nativeObjects.items, "Clone-safe ink");
const inkRoundTripClone = itemByName(inkRoundTrip.slides.getItem(1).nativeObjects.items, "Clone-safe ink");
assert.notEqual(inkRoundTripOrigin.parts[0].path, inkRoundTripClone.parts[0].path);
assert.equal(inkRoundTripOrigin.parts[0].sourceSha256, inkRoundTripClone.parts[0].sourceSha256);

const connectedInkSource = await PresentationFile.patchPptx(inkSource, [{
  path: "ppt/customXml/_rels/agent-ink.xml.rels",
  xml: '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdUnsafeInkLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/ink" TargetMode="External"/></Relationships>',
}]);
const connectedInk = await PresentationFile.importPptx(connectedInkSource);
assert.equal(connectedInk.slides.getItem(0).cloneCapability.supported, true);
const connectedInkSlideCount = connectedInk.slides.items.length;
connectedInk.slides.getItem(0).duplicate();
assert.equal((await PresentationFile.importPptx(await PresentationFile.exportPptx(connectedInk))).slides.items.length, connectedInkSlideCount + 1);

const nestedInk = await PresentationFile.importPptx(inkSource);
const nestedInkObject = itemByName(nestedInk.slides.getItem(0).nativeObjects.items, "Clone-safe ink");
nestedInkObject.rawXml = nestedInkObject.rawXml.replace(/^<p:contentPart/, "<p:grpSp");
const nestedInkSlideCount = nestedInk.slides.items.length;
assert.throws(
  () => nestedInk.slides.getItem(0).duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "an InkML graph whose source binding is not a top-level contentPart must fail before mutating the model",
);
assert.equal(nestedInk.slides.items.length, nestedInkSlideCount);

// PowerPoint represents an embedded video as one top-level picture with a
// poster ImagePart plus paired video/media data relationships to one MP4.
// The bounded clone shares the immutable poster but copies the MP4 into a new
// SDK-allocated MediaDataPart, preserving both slide-local relationship IDs.
const embeddedVideoBytes = Buffer.from("AAAAIGZ0eXBpc29tAAACAGlzb21pc28yYXZjMW1wNDEAAAMVbW9vdgAAAGxtdmhkAAAAAAAAAAAAAAAAAAAD6AAAACgAAQAAAQAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAj90cmFrAAAAXHRraGQAAAADAAAAAAAAAAAAAAABAAAAAAAAACgAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAABAAAAAABAAAAAQAAAAAAAkZWR0cwAAABxlbHN0AAAAAAAAAAEAAAAoAAAAAAABAAAAAAG3bWRpYQAAACBtZGhkAAAAAAAAAAAAAAAAAAAyAAAAAgBVxAAAAAAALWhkbHIAAAAAAAAAAHZpZGUAAAAAAAAAAAAAAABWaWRlb0hhbmRsZXIAAAABYm1pbmYAAAAUdm1oZAAAAAEAAAAAAAAAAAAAACRkaW5mAAAAHGRyZWYAAAAAAAAAAQAAAAx1cmwgAAAAAQAAASJzdGJsAAAAvnN0c2QAAAAAAAAAAQAAAK5hdmMxAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAABAAEABIAAAASAAAAAAAAAABFUxhdmM2Mi4yOC4xMDIgbGlieDI2NAAAAAAAAAAAAAAAGP//AAAANGF2Y0MBZAAK/+EAF2dkAAqs2V7ARAAAAwAEAAADAMg8SJZYAQAGaOvjyyLA/fj4AAAAABBwYXNwAAAAAQAAAAEAAAAUYnRydAAAAAAAAinoAAAAAAAAABhzdHRzAAAAAAAAAAEAAAABAAACAAAAABxzdHNjAAAAAAAAAAEAAAABAAAAAQAAAAEAAAAUc3RzegAAAAAAAALFAAAAAQAAABRzdGNvAAAAAAAAAAEAAANFAAAAYnVkdGEAAABabWV0YQAAAAAAAAAhaGRscgAAAAAAAAAAbWRpcmFwcGwAAAAAAAAAAAAAAAAtaWxzdAAAACWpdG9vAAAAHWRhdGEAAAABAAAAAExhdmY2Mi4xMi4xMDIAAAAIZnJlZQAAAs1tZGF0AAACrgYF//+q3EXpvebZSLeWLNgg2SPu73gyNjQgLSBjb3JlIDE2NSByMzIyMiBiMzU2MDVhIC0gSC4yNjQvTVBFRy00IEFWQyBjb2RlYyAtIENvcHlsZWZ0IDIwMDMtMjAyNSAtIGh0dHA6Ly93d3cudmlkZW9sYW4ub3JnL3gyNjQuaHRtbCAtIG9wdGlvbnM6IGNhYmFjPTEgcmVmPTMgZGVibG9jaz0xOjA6MCBhbmFseXNpPTB4MzoweDExMyBtZT1oZXggc3VibWU9NyBwc3k9MSBwc3lfcmQ9MS4wMDowLjAwIG1peGVkX3JlZj0xIG1lX3JhbmdlPTE2IGNocm9tYV9tZT0xIHRyZWxsaXM9MSA4eDhkY3Q9MSBjcW09MCBkZWFkem9uZT0yMSwxMSBmYXN0X3Bza2lwPTEgY2hyb21hX3FwX29mZnNldD0tMiB0aHJlYWRzPTEgbG9va2FoZWFkX3RocmVhZHM9MSBzbGljZWRfdGhyZWFkcz0wIG5yPTAgZGVjaW1hdGU9MSBpbnRlcmxhY2VkPTAgYmx1cmF5X2NvbXBhdD0wIGNvbnN0cmFpbmVkX2ludHJhPTAgYmZyYW1lcz0zIGJfcHlyYW1pZD0yIGJfYWRhcHQ9MSBiX2JpYXM9MCBkaXJlY3Q9MSB3ZWlnaHRiPTEgb3Blbl9nb3A9MCB3ZWlnaHRwPTIga2V5aW50PTI1MCBrZXlpbnRfbWluPTI1IHNjZW5lY3V0PTQwIGludHJhX3JlZnJlc2g9MCByY19sb29rYWhlYWQ9NDAgcmM9Y3JmIG1idHJlZT0xIGNyZj0yMy4wIHFjb21wPTAuNjAgcXBtaW49MCBxcG1heD02OSBxcHN0ZXA9NCBpcF9yYXRpbz0xLjQwIGFxPTE6MS4wMACAAAAAD2WIhAAr//72c3wKa22xgQ==", "base64");
const mediaPicture = '<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvPicPr><p:cNvPr id="122" name="Clone-safe video"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:videoFile r:link="rIdAgentVideo"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media r:embed="rIdAgentMedia"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="rIdAgentVideoPoster"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="3657600" cy="2286000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
const mediaRelationships = '<Relationship Id="rIdAgentVideo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/agent-video.mp4"/><Relationship Id="rIdAgentMedia" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/agent-video.mp4"/><Relationship Id="rIdAgentVideoPoster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/agent-video-poster.png"/>';
const mediaSource = await PresentationFile.patchPptx(cloneSourcePptx, [
  { path: "ppt/slides/slide1.xml", xml: oleCloneBaseSlideXml.replace("</p:spTree>", `${mediaPicture}</p:spTree>`) },
  { path: "ppt/slides/_rels/slide1.xml.rels", xml: oleCloneBaseRelationships.replace("</Relationships>", `${mediaRelationships}</Relationships>`) },
  { path: "ppt/media/agent-video.mp4", bytes: embeddedVideoBytes, contentType: "video/mp4" },
  { path: "ppt/media/agent-video-poster.png", bytes: embeddedPreviewBytes, contentType: "image/png" },
]);
const mediaSourceSnapshot = Uint8Array.from(mediaSource.bytes);
const mediaImported = await PresentationFile.importPptx(mediaSource);
const mediaOriginSlide = mediaImported.slides.getItem(0);
const mediaOrigin = itemByName(mediaOriginSlide.nativeObjects.items, "Clone-safe video");
assert.equal(mediaOrigin.nativeKind, "media");
assert.equal(mediaOrigin.relationshipReferences.length, 3);
assert.equal(mediaOrigin.parts.length, 2);
assert.equal(mediaOrigin.parts.filter((part) => part.contentType === "video/mp4").length, 1);
assert.equal(mediaOrigin.parts.filter((part) => part.contentType === "image/png").length, 1);
const mediaPendingSlide = mediaOriginSlide.duplicate();
const mediaPending = itemByName(mediaPendingSlide.nativeObjects.items, "Clone-safe video");
assert.notEqual(mediaPending, mediaOrigin);
assert.notEqual(mediaPending.id, mediaOrigin.id);
assert.deepEqual(mediaPending.parts.map((part) => part.path), mediaOrigin.parts.map((part) => part.path));

const mediaExport = await PresentationFile.exportPptx(mediaImported);
assert.deepEqual(mediaSource.bytes, mediaSourceSnapshot);
const mediaSourceZip = await JSZip.loadAsync(mediaSource.bytes);
const mediaOutputZip = await JSZip.loadAsync(mediaExport.bytes);
const mediaCloneSlidePath = (await orderedPptxSlidePaths(mediaOutputZip))[1];
assert.deepEqual(
  await mediaOutputZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await mediaSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "embedded-video cloning must retain the origin SlidePart byte-for-byte",
);
assert.deepEqual(
  await mediaOutputZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  await mediaSourceZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  "embedded-video cloning must retain the origin relationship part byte-for-byte",
);
const mediaSourceRels = await mediaOutputZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const mediaCloneRels = await mediaOutputZip.file(relationshipPartPath(mediaCloneSlidePath)).async("text");
const sourceVideoRelationship = relationshipForType(mediaSourceRels, "video");
const sourceMediaRelationship = relationshipForType(mediaSourceRels, "media");
const sourcePosterRelationship = relationshipForType(mediaSourceRels, "image");
const cloneVideoRelationship = relationshipForType(mediaCloneRels, "video");
const cloneMediaRelationship = relationshipForType(mediaCloneRels, "media");
const clonePosterRelationship = relationshipForType(mediaCloneRels, "image");
assert.equal(cloneVideoRelationship.id, sourceVideoRelationship.id);
assert.equal(cloneMediaRelationship.id, sourceMediaRelationship.id);
assert.equal(clonePosterRelationship.id, sourcePosterRelationship.id);
const sourceVideoPartPath = resolveSlideRelationshipTarget(sourceVideoRelationship.target);
const sourceMediaPartPath = resolveSlideRelationshipTarget(sourceMediaRelationship.target);
const cloneVideoPartPath = resolveSlideRelationshipTarget(cloneVideoRelationship.target);
const cloneMediaPartPath = resolveSlideRelationshipTarget(cloneMediaRelationship.target);
assert.equal(sourceVideoPartPath, sourceMediaPartPath);
assert.equal(cloneVideoPartPath, cloneMediaPartPath);
assert.match(cloneVideoPartPath, /^(?:ppt\/)?media\/[^/]+\.mp4$/i);
assert.notEqual(cloneVideoPartPath, sourceVideoPartPath);
assert.deepEqual(
  await mediaOutputZip.file(cloneVideoPartPath).async("uint8array"),
  await mediaSourceZip.file(sourceVideoPartPath).async("uint8array"),
  "embedded-video cloning must byte-copy the accepted MP4 into an independent MediaDataPart",
);
assert.equal(
  resolveSlideRelationshipTarget(clonePosterRelationship.target),
  resolveSlideRelationshipTarget(sourcePosterRelationship.target),
  "embedded-video cloning must share the immutable poster ImagePart",
);

const mediaRoundTrip = await PresentationFile.importPptx(mediaExport);
const mediaRoundTripOrigin = itemByName(mediaRoundTrip.slides.getItem(0).nativeObjects.items, "Clone-safe video");
const mediaRoundTripClone = itemByName(mediaRoundTrip.slides.getItem(1).nativeObjects.items, "Clone-safe video");
const mediaRoundTripOriginVideo = mediaRoundTripOrigin.parts.find((part) => part.contentType === "video/mp4");
const mediaRoundTripCloneVideo = mediaRoundTripClone.parts.find((part) => part.contentType === "video/mp4");
const mediaRoundTripOriginPoster = mediaRoundTripOrigin.parts.find((part) => part.contentType.startsWith("image/"));
const mediaRoundTripClonePoster = mediaRoundTripClone.parts.find((part) => part.contentType.startsWith("image/"));
assert.notEqual(mediaRoundTripOriginVideo.path, mediaRoundTripCloneVideo.path);
assert.equal(mediaRoundTripOriginVideo.sourceSha256, mediaRoundTripCloneVideo.sourceSha256);
assert.equal(mediaRoundTripOriginPoster.path, mediaRoundTripClonePoster.path);

const malformedMedia = await PresentationFile.importPptx(mediaSource);
const malformedMediaObject = itemByName(malformedMedia.slides.getItem(0).nativeObjects.items, "Clone-safe video");
malformedMediaObject.rawXml = malformedMediaObject.rawXml.replace(
  "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}",
  "{00000000-0000-0000-0000-000000000000}",
);
const malformedMediaSlideCount = malformedMedia.slides.items.length;
assert.throws(
  () => malformedMedia.slides.getItem(0).duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "a media picture with a non-canonical extension must fail before mutating the model",
);
assert.equal(malformedMedia.slides.items.length, malformedMediaSlideCount);

const extensionRichMedia = await PresentationFile.importPptx(mediaSource);
const extensionRichMediaObject = itemByName(extensionRichMedia.slides.getItem(0).nativeObjects.items, "Clone-safe video");
extensionRichMediaObject.rawXml = extensionRichMediaObject.rawXml.replace(
  "</p:extLst>",
  '<p:ext uri="{00000000-0000-0000-0000-000000000000}"><p14:placeholder/></p:ext></p:extLst>',
);
const extensionRichMediaSlideCount = extensionRichMedia.slides.items.length;
assert.throws(
  () => extensionRichMedia.slides.getItem(0).duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "a media picture with an extra extension must fail before mutating the model",
);
assert.equal(extensionRichMedia.slides.items.length, extensionRichMediaSlideCount);

const wrongTypeMediaSource = await PresentationFile.patchPptx(cloneSourcePptx, [
  { path: "ppt/slides/slide1.xml", xml: oleCloneBaseSlideXml.replace("</p:spTree>", `${mediaPicture}</p:spTree>`) },
  { path: "ppt/slides/_rels/slide1.xml.rels", xml: oleCloneBaseRelationships.replace("</Relationships>", `${mediaRelationships}</Relationships>`) },
  { path: "ppt/media/agent-video.mp4", bytes: embeddedVideoBytes, contentType: "video/quicktime" },
  { path: "ppt/media/agent-video-poster.png", bytes: embeddedPreviewBytes, contentType: "image/png" },
]);
const wrongTypeMedia = await PresentationFile.importPptx(wrongTypeMediaSource);
assert.equal(wrongTypeMedia.slides.getItem(0).cloneCapability.supported, true);
const wrongTypeMediaSlideCount = wrongTypeMedia.slides.items.length;
wrongTypeMedia.slides.getItem(0).duplicate();
assert.equal((await PresentationFile.importPptx(await PresentationFile.exportPptx(wrongTypeMedia))).slides.items.length, wrongTypeMediaSlideCount + 1);

const masterPath = "ppt/slideMasters/slideMaster1.xml";
const layoutPath = "ppt/slideLayouts/slideLayout1.xml";
const masterXml = await firstZip.file(masterPath).async("text");
const layoutXml = await firstZip.file(layoutPath).async("text");
const sourceMasterXml = masterXml.replace(/(<p:cSld\b[^>]*\bname=")[^"]*(")/, "$1Source Master Marker$2");
const sourceLayoutXml = layoutXml.replace(/(<p:cSld\b[^>]*\bname=")[^"]*(")/, "$1Source Layout Marker$2");
assert.notEqual(sourceMasterXml, masterXml);
assert.notEqual(sourceLayoutXml, layoutXml);
const sourceBound = await PresentationFile.patchPptx(firstExport, [
  { path: masterPath, xml: sourceMasterXml },
  { path: layoutPath, xml: sourceLayoutXml },
]);

const imported = await PresentationFile.importPptx(sourceBound);
assert.equal(imported.master.name, "Source Master Marker");
assert.equal(imported.layouts.items[0].name, "Source Layout Marker");
imported.master.name = "Unsupported master edit";
await assert.rejects(
  () => PresentationFile.exportPptx(imported),
  /master .*source-bound and read-only/i,
);
imported.master.name = "Source Master Marker";
imported.layouts.items[0].name = "Unsupported layout edit";
await assert.rejects(
  () => PresentationFile.exportPptx(imported),
  /layout .*source-bound and read-only/i,
);
imported.layouts.items[0].name = "Source Layout Marker";
assert.equal(imported.slides.getItem(0).speakerNotes.text, "Lead with the customer outcome.\nThen explain the operating model.");
assert.deepEqual(imported.slides.getItem(0).speakerNotes.capability, {
  sourceBound: true,
  partPresent: true,
  editable: true,
  addable: false,
});
assert.deepEqual(imported.slides.getItem(1).speakerNotes.capability, {
  sourceBound: true,
  partPresent: false,
  editable: false,
  addable: true,
});
imported.slides.getItem(0).addNotes("Lead with evidence.\nClose with the decision.");
imported.slides.getItem(1).addNotes("Explain the chart assumptions.\nInvite questions on the forecast.");
const importedCore = imported.slides.getItem(0);
assert.deepEqual(importedCore.background, { fill: "#f1f5f9", mode: "solid" });
assert.equal(itemByName(importedCore.shapes.items, "rounded-card").geometry, "roundRect");
assert.equal(itemByName(importedCore.shapes.items, "target-textbox").geometry, "textbox");
assert.deepEqual(itemByName(importedCore.shapes.items, "rounded-card").shadow, {
  color: "#000000",
  blurRadius: 8,
  distance: 4,
  direction: 45,
  opacity: 0.25,
});
assert.equal(itemByName(importedCore.images.items, "png-image").dataUrl, PNG);
assert.equal(itemByName(importedCore.images.items, "jpeg-image").dataUrl, JPEG);
const importedCover = itemByName(importedCore.images.items, "cover-image");
assert.equal(importedCover.fit, "stretch");
assert.deepEqual(importedCover.crop, { left: 0.25, top: 0, right: 0.25, bottom: 0 });
assert.equal(itemByName(importedCore.tables.items, "fixed-table").values[1][1], "Before");
const importedStraight = itemByName(importedCore.connectors.items, "straight-connector");
const importedElbow = itemByName(importedCore.connectors.items, "elbow-polyline-connector");
assert.equal(importedStraight.line.endArrow, "triangle");
assert.equal(importedElbow.connectorType, "elbow");
assert.equal(importedElbow.line.startArrow, "triangle");
assert.equal(importedElbow.line.endArrow, "triangle");
assert.ok(importedElbow.startTargetId && importedElbow.endTargetId);
const importedRich = itemByName(importedCore.shapes.items, "rich-copy");
assert.equal(importedRich.text.paragraphs[1].bulletCharacter, "•");
assert.deepEqual(importedRich.text.paragraphs[0].runs[1].link, {
  uri: "https://www.ecma-international.org/publications-and-standards/standards/ecma-376/",
});
const importedCharts = imported.slides.getItem(1).charts.items;
assert.deepEqual(importedCharts.map((chart) => chart.chartType), ["bar", "line", "pie"]);
assert.equal(importedCharts[1].series[0].marker.symbol, "circle");
assert.deepEqual(importedCharts[1].series[0].trendlines.map((trendline) => trendline.type), ["linear", "movingAvg", "poly"]);
assert.equal(importedCharts[1].series[0].trendlines[0].name, "Pass projection");
assert.equal(importedCharts[1].series[0].trendlines[0].displayRSquared, true);
assert.equal(importedCharts[1].series[0].trendlines[0].line.fill, "#7C3AED");
assert.deepEqual(importedCharts[1].series[0].errorBars, {
  direction: "y",
  type: "both",
  valueType: "stdDev",
  value: 1.5,
  noEndCap: true,
  line: { fill: "#DC2626", style: "dot", width: 1.25 },
});
assert.equal(importedCharts[2].dataLabels.showCategoryName, true);
const importedCombo = itemByName(imported.slides.getItem(2).charts.items, "revenue-margin-combo");
assert.equal(importedCombo.chartType, "combo");
assert.deepEqual(importedCombo.series.map((series) => series.chartType), ["bar", "line"]);
assert.equal(importedCombo.series[1].marker.symbol, "circle");
assert.equal(importedCombo.series[1].trendlines[0].type, "exp");
assert.equal(importedCombo.series[1].trendlines[0].name, "Margin projection");
assert.deepEqual(importedCombo.series[1].errorBars, {
  direction: "y",
  type: "minus",
  valueType: "cust",
  minusValues: [1, 2, 1],
  noEndCap: false,
  line: { fill: "#EA580C", style: "solid", width: 1 },
});
assert.equal(importedCombo.dataLabels.showValue, true);
assert.equal(importedCombo.dataLabels.position, "t");

const importedCard = itemByName(importedCore.shapes.items, "rounded-card");
importedCard.text.set("After edit");
importedCard.shadow.opacity = 0.35;
assert.equal(importedCore.setBackground({ fill: "accent2", mode: "reference", index: 1002 }), importedCore);
assert.equal(imported.slides.getItem(1).setBackground({ fill: "#FFF7ED", mode: "solid" }), imported.slides.getItem(1));
itemByName(importedCore.tables.items, "fixed-table").cells.set(1, 1, "After");
itemByName(importedCore.images.items, "png-image").alt = "Updated PNG evidence";
importedCover.fit = "contain";
importedCover.crop = undefined;
delete importedElbow.line.endArrow;
const editedParagraphs = importedRich.text.paragraphs;
editedParagraphs[0].runs[0].text = "Updated ";
importedRich.text.paragraphs = editedParagraphs;
const importedBar = itemByName(importedCharts, "bar-chart");
importedBar.title = "Updated readiness";
importedBar.series[0].values = [80, 94, 88];
importedCharts[1].series[0].trendlines[0].name = "Edited pass projection";
importedCharts[1].series[0].trendlines[0].forward = 1.5;
importedCharts[1].series[0].trendlines[0].line.fill = "#0EA5E9";
importedCharts[1].series[0].errorBars.value = 2;
importedCharts[1].series[0].errorBars.line.fill = "#BE123C";
importedCombo.title = "Updated revenue and margin";
importedCombo.series[1].values = [12, 16, 18];
importedCombo.series[1].trendlines[0].name = "Edited margin projection";
importedCombo.series[1].errorBars.minusValues[1] = 3;

const secondExport = await PresentationFile.exportPptx(imported);
assert.equal(secondExport.metadata.codec, "office-kit");
assert.equal((await PresentationFile.inspectPptx(secondExport)).ok, true);
const secondZip = await JSZip.loadAsync(new Uint8Array(await secondExport.arrayBuffer()));
assert.equal(await secondZip.file(masterPath).async("text"), sourceMasterXml);
assert.equal(await secondZip.file(layoutPath).async("text"), sourceLayoutXml);
assert.match(await secondZip.file("ppt/slides/_rels/slide1.xml.rels").async("text"), /relationships\/slideLayout/);
assert.match(await secondZip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels").async("text"), /relationships\/slideMaster/);
const secondSlideXml = await secondZip.file("ppt/slides/slide1.xml").async("text");
const secondChartSlideXml = await secondZip.file("ppt/slides/slide2.xml").async("text");
assert.match(secondSlideXml, /<p:bgRef idx="1002">/);
assert.match(secondSlideXml, /<a:schemeClr val="accent2"/);
assert.match(secondChartSlideXml, /<a:srgbClr val="FFF7ED"/);
assert.match(secondSlideXml, /<a:srcRect[^>]*t="-50000"/);
assert.match(secondSlideXml, /<a:srcRect[^>]*b="-50000"/);
assert.match(secondSlideXml, /prst="roundRect"/);
assert.match(secondSlideXml, /txBox="1"/);
assert.match(secondSlideXml, /prst="straightConnector1"/);
assert.match(secondSlideXml, /prst="bentConnector3"/);
assert.match(secondSlideXml, /<a:headEnd type="triangle"/);
assert.match(secondSlideXml, /<a:tailEnd type="triangle"/);
assert.ok(Object.keys(secondZip.files).some((name) => /\/media\/.+\.png$/.test(name)));
assert.ok(Object.keys(secondZip.files).some((name) => /\/media\/.+\.jpe?g$/.test(name)));
assert.equal(Object.keys(secondZip.files).filter((name) => /\/charts\/chart\d+\.xml$/.test(name)).length, 4);
const secondChartXml = await Promise.all(Object.keys(secondZip.files)
  .filter((name) => /\/charts\/chart\d+\.xml$/.test(name))
  .map((name) => secondZip.file(name).async("text")));
assert.ok(secondChartXml.some((xml) => xml.includes("Edited pass projection") && /<c:forward val="1\.5"\s*\/>/.test(xml)));
assert.ok(secondChartXml.some((xml) => xml.includes("Edited margin projection") && /<c:trendlineType val="exp"\s*\/>/.test(xml)));
assert.ok(secondChartXml.some((xml) => /<c:errValType val="stdDev"\s*\/>/.test(xml) && /<c:val val="2"\s*\/>/.test(xml)));
assert.ok(secondChartXml.some((xml) => /<c:errBarType val="minus"\s*\/>/.test(xml) && />3<\//.test(xml)));
assert.match(await secondZip.file("ppt/notesSlides/notesSlide1.xml").async("text"), /Lead with evidence/);
assert.equal(Object.keys(secondZip.files).filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name)).length, 2);
assert.ok((await Promise.all(
  Object.keys(secondZip.files)
    .filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name))
    .map(async (name) => (await secondZip.file(name).async("text")).includes("Explain the chart assumptions")),
)).includes(true));

const roundTrip = await PresentationFile.importPptx(secondExport);
assert.equal(roundTrip.master.name, "Source Master Marker");
assert.equal(roundTrip.layouts.items[0].name, "Source Layout Marker");
const roundTripCore = roundTrip.slides.getItem(0);
assert.equal(roundTripCore.speakerNotes.text, "Lead with evidence.\nClose with the decision.");
assert.equal(roundTrip.slides.getItem(1).speakerNotes.text, "Explain the chart assumptions.\nInvite questions on the forecast.");
assert.deepEqual(roundTrip.slides.getItem(1).speakerNotes.capability, {
  sourceBound: true,
  partPresent: true,
  editable: true,
  addable: false,
});
assert.deepEqual(roundTripCore.background, { fill: "accent2", mode: "reference", index: 1002 });
assert.deepEqual(roundTrip.slides.getItem(1).background, { fill: "#fff7ed", mode: "solid" });
assert.equal(itemByName(roundTripCore.shapes.items, "rounded-card").text.value, "After edit");
assert.equal(itemByName(roundTripCore.shapes.items, "rounded-card").shadow.opacity, 0.35);
assert.equal(itemByName(roundTripCore.tables.items, "fixed-table").values[1][1], "After");
assert.equal(itemByName(roundTripCore.images.items, "png-image").alt, "Updated PNG evidence");
const roundTripCover = itemByName(roundTripCore.images.items, "cover-image");
assert.equal(roundTripCover.fit, "stretch");
assert.deepEqual(roundTripCover.crop, { left: 0, top: -0.5, right: 0, bottom: -0.5 });
assert.equal(itemByName(roundTripCore.connectors.items, "elbow-polyline-connector").line.endArrow, undefined);
assert.equal(itemByName(roundTripCore.shapes.items, "rich-copy").text.paragraphs[0].runs[0].text, "Updated ");
const roundTripBar = itemByName(roundTrip.slides.getItem(1).charts.items, "bar-chart");
assert.equal(roundTripBar.title, "Updated readiness");
assert.deepEqual(roundTripBar.series[0].values, [80, 94, 88]);
const roundTripLine = itemByName(roundTrip.slides.getItem(1).charts.items, "line-chart");
assert.equal(roundTripLine.series[0].trendlines[0].name, "Edited pass projection");
assert.equal(roundTripLine.series[0].trendlines[0].forward, 1.5);
assert.equal(roundTripLine.series[0].trendlines[0].line.fill, "#0EA5E9");
assert.equal(roundTripLine.series[0].errorBars.value, 2);
assert.equal(roundTripLine.series[0].errorBars.line.fill, "#BE123C");
const roundTripCombo = itemByName(roundTrip.slides.getItem(2).charts.items, "revenue-margin-combo");
assert.equal(roundTripCombo.title, "Updated revenue and margin");
assert.deepEqual(roundTripCombo.series.map((series) => series.chartType), ["bar", "line"]);
assert.deepEqual(roundTripCombo.series[1].values, [12, 16, 18]);
assert.equal(roundTripCombo.series[1].trendlines[0].name, "Edited margin projection");
assert.deepEqual(roundTripCombo.series[1].errorBars.minusValues, [1, 3, 1]);
assert.equal(roundTrip.verify().ok, true);

assert.equal(roundTripCore.clearBackground(), roundTripCore);
roundTripCover.fit = "stretch";
roundTripCover.crop = undefined;
const clearedBackgroundExport = await PresentationFile.exportPptx(roundTrip);
const clearedBackgroundZip = await JSZip.loadAsync(new Uint8Array(await clearedBackgroundExport.arrayBuffer()));
assert.doesNotMatch(await clearedBackgroundZip.file("ppt/slides/slide1.xml").async("text"), /<p:bg(?:Pr|Ref)\b/);
assert.doesNotMatch(await clearedBackgroundZip.file("ppt/slides/slide1.xml").async("text"), /<a:srcRect\b/);
const clearedBackgroundRoundTrip = await PresentationFile.importPptx(clearedBackgroundExport);
assert.deepEqual(clearedBackgroundRoundTrip.slides.getItem(0).background, {});
assert.equal(itemByName(clearedBackgroundRoundTrip.slides.getItem(0).images.items, "cover-image").crop, undefined);

const importedWithoutSourceSnapshot = await PresentationFile.importPptx(firstExport);
const presentationState = importedWithoutSourceSnapshot[Symbol.for("office-kit.presentation-state")];
presentationState.opaqueOpc.sourcePackage = undefined;
await assert.rejects(
  () => PresentationFile.exportPptx(importedWithoutSourceSnapshot),
  (error) => error?.code === "missing_source_package",
);

// OfficeKit owns a deliberately narrow legacy PPTX comment profile: one
// slide-level text item at an explicit coordinate. It never turns the richer
// JS thread facade into a fake element anchor, reply graph, or resolved state.
const legacyAdditionSourceDeck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const legacyAdditionTarget = legacyAdditionSourceDeck.slides.add({ name: "Imported review target" });
legacyAdditionTarget.shapes.add({
  name: "visible-review-title",
  geometry: "textbox",
  text: "Visible content must not change",
  position: { left: 96, top: 96, width: 900, height: 88 },
});
const legacyAdditionControl = legacyAdditionSourceDeck.slides.add({ name: "Imported review control" });
legacyAdditionControl.shapes.add({
  name: "visible-control-title",
  geometry: "textbox",
  text: "Control slide",
  position: { left: 96, top: 96, width: 900, height: 88 },
});
const legacyAdditionSource = await PresentationFile.exportPptx(legacyAdditionSourceDeck);
const legacyAdditionSourceBytes = new Uint8Array(await legacyAdditionSource.arrayBuffer());
const legacyAdditionImported = await PresentationFile.importPptx(legacyAdditionSource);
assert.deepEqual(legacyAdditionImported.slides.getItem(0).comments.capability, {
  sourceBound: true,
  format: "legacy",
  partPresent: false,
  editable: false,
  addable: true,
});
assert.match(legacyAdditionImported.inspect({ kind: "slide" }).ndjson, /"commentsCapability":\{"sourceBound":true,"format":"legacy","partPresent":false,"editable":false,"addable":true\}/);
legacyAdditionImported.slides.getItem(0).comments.addThread(undefined, "Confirm the imported evidence.", {
  author: "Review Owner",
  created: "2026-07-20T03:04:05Z",
  position: { x: 360, y: 240 },
});
const legacyAdditionExport = await PresentationFile.exportPptx(legacyAdditionImported);
const legacyAdditionOutputBytes = new Uint8Array(await legacyAdditionExport.arrayBuffer());
const legacyAdditionSourceZip = await JSZip.loadAsync(legacyAdditionSourceBytes);
const legacyAdditionOutputZip = await JSZip.loadAsync(legacyAdditionOutputBytes);
assert.deepEqual(
  await legacyAdditionOutputZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await legacyAdditionSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
);
assert.deepEqual(
  await legacyAdditionOutputZip.file("ppt/slides/slide2.xml").async("uint8array"),
  await legacyAdditionSourceZip.file("ppt/slides/slide2.xml").async("uint8array"),
);
assert.ok(legacyAdditionOutputZip.file("ppt/commentAuthors.xml"));
assert.ok(legacyAdditionOutputZip.file("ppt/comments/comment1.xml"));
const legacyAdditionRoundTrip = await PresentationFile.importPptx(legacyAdditionExport);
assert.equal(legacyAdditionRoundTrip.slides.getItem(0).comments.items[0].comments[0].text, "Confirm the imported evidence.");
assert.deepEqual(legacyAdditionRoundTrip.slides.getItem(0).comments.capability, {
  sourceBound: true,
  format: "legacy",
  partPresent: true,
  editable: true,
  addable: false,
});
assert.deepEqual(legacyAdditionRoundTrip.slides.getItem(1).comments.capability, {
  sourceBound: true,
  format: "legacy",
  partPresent: false,
  editable: false,
  addable: false,
});

const legacyCommentDeck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const legacyCommentSlide = legacyCommentDeck.slides.add({ name: "Legacy comments" });
const legacyCommentThread = legacyCommentSlide.comments.addThread(undefined, "Confirm the source before delivery.", {
  author: "Review Owner",
  created: "2026-07-18T03:05:00Z",
  position: { x: 360, y: 240 },
});
assert.match(legacyCommentDeck.inspect({ kind: "comment" }).ndjson, /Confirm the source before delivery/);
const legacyCommentExport = await PresentationFile.exportPptx(legacyCommentDeck);
const legacyCommentZip = await JSZip.loadAsync(new Uint8Array(await legacyCommentExport.arrayBuffer()));
assert.ok(legacyCommentZip.file("ppt/comments/comment1.xml"));
assert.ok(legacyCommentZip.file("ppt/commentAuthors.xml"));
assert.match(await legacyCommentZip.file("ppt/comments/comment1.xml").async("text"), /Confirm the source before delivery/);
const legacyCommentImported = await PresentationFile.importPptx(legacyCommentExport);
assert.equal(legacyCommentImported.slides.getItem(0).comments.items.length, 1);
assert.deepEqual(legacyCommentImported.slides.getItem(0).comments.capability, {
  sourceBound: true,
  format: "legacy",
  partPresent: true,
  editable: true,
  addable: false,
});
const importedLegacyThread = legacyCommentImported.slides.getItem(0).comments.items[0];
assert.equal(importedLegacyThread.nativeFormat, "legacy");
assert.equal(importedLegacyThread.targetId, undefined);
assert.equal(importedLegacyThread.comments.length, 1);
assert.equal(importedLegacyThread.comments[0].author, "Review Owner");
assert.equal(importedLegacyThread.comments[0].text, "Confirm the source before delivery.");
assert.deepEqual(importedLegacyThread.position, { x: 360, y: 240, unit: "px" });

// The bounded imported-slide clone profile may carry a closed legacy-comments
// leaf. The clone gets a distinct model/thread object and comments part, while
// both p:cm records keep their IDs against the one immutable author catalog.
const legacyCommentCloneDeck = await PresentationFile.importPptx(legacyCommentExport);
const legacyCommentCloneSource = legacyCommentCloneDeck.slides.getItem(0);
const legacyCommentClone = legacyCommentCloneSource.duplicate();
assert.equal(legacyCommentClone.comments.items.length, 1);
assert.notEqual(legacyCommentClone.comments.items[0], legacyCommentCloneSource.comments.items[0]);
assert.equal(legacyCommentClone.comments.items[0].comments[0].text, "Confirm the source before delivery.");
const legacyCommentCloneExport = await PresentationFile.exportPptx(legacyCommentCloneDeck);
const legacyCommentCloneZip = await JSZip.loadAsync(new Uint8Array(await legacyCommentCloneExport.arrayBuffer()));
const legacyCommentClonePaths = Object.keys(legacyCommentCloneZip.files).filter((partPath) => /^ppt\/comments\/comment[^/]*\.xml$/i.test(partPath));
assert.equal(legacyCommentClonePaths.length, 2);
const legacyCommentClonePartPath = legacyCommentClonePaths.find((partPath) => partPath !== "ppt/comments/comment1.xml");
assert.ok(legacyCommentClonePartPath);
assert.deepEqual(
  await legacyCommentCloneZip.file(legacyCommentClonePartPath).async("uint8array"),
  await legacyCommentZip.file("ppt/comments/comment1.xml").async("uint8array"),
);
assert.deepEqual(
  await legacyCommentCloneZip.file("ppt/comments/comment1.xml").async("uint8array"),
  await legacyCommentZip.file("ppt/comments/comment1.xml").async("uint8array"),
);
assert.deepEqual(
  await legacyCommentCloneZip.file("ppt/commentAuthors.xml").async("uint8array"),
  await legacyCommentZip.file("ppt/commentAuthors.xml").async("uint8array"),
);
const legacyCommentCloneRoundTrip = await PresentationFile.importPptx(legacyCommentCloneExport);
assert.equal(legacyCommentCloneRoundTrip.slides.items.length, 2);
assert.deepEqual(
  legacyCommentCloneRoundTrip.slides.items.map((slide) => slide.comments.items[0].comments[0].text),
  ["Confirm the source before delivery.", "Confirm the source before delivery."],
);

const editedLegacyCommentCloneDeck = await PresentationFile.importPptx(legacyCommentExport);
const editedLegacyCommentClone = editedLegacyCommentCloneDeck.slides.getItem(0).duplicate();
editedLegacyCommentClone.comments.items[0].comments[0].text = "This comment cannot change before the clone boundary.";
await assert.rejects(
  () => PresentationFile.exportPptx(editedLegacyCommentCloneDeck),
  (error) => error?.code === "unsupported_presentation_edit",
);

const legacyCommentRoundTrip = await PresentationFile.exportPptx(legacyCommentImported);
const legacyCommentRoundTripZip = await JSZip.loadAsync(new Uint8Array(await legacyCommentRoundTrip.arrayBuffer()));
assert.equal(
  await legacyCommentRoundTripZip.file("ppt/comments/comment1.xml").async("text"),
  await legacyCommentZip.file("ppt/comments/comment1.xml").async("text"),
);

// A recognized imported legacy leaf exposes exactly one safe mutation: the
// root text. Its author catalog, native comment identity, coordinate, SlidePart
// XML, and relationship graph must remain byte-stable.
importedLegacyThread.comments[0].text = "Confirm the source and attach the delivery evidence.";
const legacyCommentTextEdit = await PresentationFile.exportPptx(legacyCommentImported);
const legacyCommentTextEditZip = await JSZip.loadAsync(new Uint8Array(await legacyCommentTextEdit.arrayBuffer()));
assert.notEqual(
  await legacyCommentTextEditZip.file("ppt/comments/comment1.xml").async("text"),
  await legacyCommentZip.file("ppt/comments/comment1.xml").async("text"),
);
for (const path of ["ppt/commentAuthors.xml", "ppt/slides/slide1.xml", "ppt/slides/_rels/slide1.xml.rels", "[Content_Types].xml"]) {
  assert.deepEqual(
    await legacyCommentTextEditZip.file(path).async("uint8array"),
    await legacyCommentZip.file(path).async("uint8array"),
    `legacy comment text edit must leave ${path} byte-identical`,
  );
}
const legacyCommentTextEditRoundTrip = await PresentationFile.importPptx(legacyCommentTextEdit);
const editedLegacyThread = legacyCommentTextEditRoundTrip.slides.getItem(0).comments.items[0];
assert.equal(editedLegacyThread.comments[0].text, "Confirm the source and attach the delivery evidence.");
assert.equal(editedLegacyThread.comments[0].author, "Review Owner");
assert.equal(editedLegacyThread.created, importedLegacyThread.created);
assert.deepEqual(editedLegacyThread.position, { x: 360, y: 240, unit: "px" });
assert.deepEqual(legacyCommentTextEditRoundTrip.slides.getItem(0).comments.capability, {
  sourceBound: true,
  format: "legacy",
  partPresent: true,
  editable: true,
  addable: false,
});

// The JavaScript boundary must preserve every non-target comment in the same
// native comments leaf; the C# profile then applies only the requested text
// payload while retaining both package-local indexes.
const multiLegacyCommentDeck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const multiLegacyCommentSlide = multiLegacyCommentDeck.slides.add({ name: "Multiple legacy comments" });
multiLegacyCommentSlide.comments.addThread(undefined, "Keep this first review note unchanged.", {
  author: "Review Owner",
  created: "2026-07-18T03:05:00Z",
  position: { x: 240, y: 160 },
});
multiLegacyCommentSlide.comments.addThread(undefined, "Replace only this second review note.", {
  author: "Review Owner",
  created: "2026-07-18T03:06:00Z",
  position: { x: 480, y: 320 },
});
const multiLegacyCommentSource = await PresentationFile.exportPptx(multiLegacyCommentDeck);
const multiLegacyCommentImported = await PresentationFile.importPptx(multiLegacyCommentSource);
assert.equal(multiLegacyCommentImported.slides.getItem(0).comments.capability.editable, true);
multiLegacyCommentImported.slides.getItem(0).comments.items[1].comments[0].text = "The second review note has the approved wording.";
const multiLegacyCommentEdited = await PresentationFile.exportPptx(multiLegacyCommentImported);
const multiLegacyCommentRoundTrip = await PresentationFile.importPptx(multiLegacyCommentEdited);
assert.deepEqual(
  multiLegacyCommentRoundTrip.slides.getItem(0).comments.items.map((thread) => ({
    text: thread.comments[0].text,
    position: thread.position,
    nativeIndex: thread.nativeAnchor.nativeIndex,
  })),
  [
    { text: "Keep this first review note unchanged.", position: { x: 240, y: 160, unit: "px" }, nativeIndex: 1 },
    { text: "The second review note has the approved wording.", position: { x: 480, y: 320, unit: "px" }, nativeIndex: 2 },
  ],
);

const legacyCommentPositionMutation = await PresentationFile.importPptx(legacyCommentExport);
legacyCommentPositionMutation.slides.getItem(0).comments.items[0].position.x += 1;
await assert.rejects(
  () => PresentationFile.exportPptx(legacyCommentPositionMutation),
  (error) => error?.code === "unsupported_presentation_edit",
);
const legacyCommentAuthorMutation = await PresentationFile.importPptx(legacyCommentExport);
legacyCommentAuthorMutation.slides.getItem(0).comments.items[0].comments[0].author = "Different reviewer";
await assert.rejects(
  () => PresentationFile.exportPptx(legacyCommentAuthorMutation),
  (error) => error?.code === "unsupported_presentation_edit",
);
const legacyCommentNativeIdentityMutation = await PresentationFile.importPptx(legacyCommentExport);
legacyCommentNativeIdentityMutation.slides.getItem(0).comments.items[0].nativeAnchor.nativeIndex += 1;
await assert.rejects(
  () => PresentationFile.exportPptx(legacyCommentNativeIdentityMutation),
  (error) => error?.code === "unsupported_presentation_edit",
);
importedLegacyThread.addReply("Replies are not part of the legacy profile.");
await assert.rejects(
  () => PresentationFile.exportPptx(legacyCommentImported),
  (error) => error?.code === "unsupported_presentation_edit",
);

const invalidLegacyCommentDeck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const invalidLegacyCommentSlide = invalidLegacyCommentDeck.slides.add();
const invalidLegacyTarget = invalidLegacyCommentSlide.shapes.add({
  geometry: "rect",
  position: { left: 40, top: 40, width: 160, height: 80 },
});
invalidLegacyCommentSlide.comments.addThread(invalidLegacyTarget, "An element anchor is not a legacy comment.", {
  author: "Reviewer",
  position: { x: 120, y: 80 },
});
await assert.rejects(
  () => PresentationFile.exportPptx(invalidLegacyCommentDeck),
  (error) => error?.code === "unsupported_presentation_features",
);
assert.equal(legacyCommentThread.id.startsWith("pc"), true);

// Office 2021 modern comments use their native author/comments graph instead
// of the legacy slide annotation part. OfficeKit owns a bounded root +
// direct replies profile with top-level drawing and shape-text-range anchors.
const modernCommentDeck = Presentation.create({
  slideSize: { width: 1280, height: 720 },
  commentFormat: "modern",
});
const modernCommentSlide = modernCommentDeck.slides.add({ name: "Modern comments" });
const modernCommentTarget = modernCommentSlide.shapes.add({
  id: "modern-comment-target",
  name: "Decision evidence",
  geometry: "rect",
  position: { left: 80, top: 80, width: 520, height: 120 },
  text: "Customer evidence is ready",
});
const modernCommentThread = modernCommentSlide.comments.addThread({
  textMatch: { element: modernCommentTarget, query: "Customer evidence is ready", occurrence: 0 },
}, "Confirm the customer evidence.", {
  id: "{11111111-1111-4111-8111-111111111111}",
  author: "Review Owner",
  created: "2026-07-19T02:55:00Z",
  nativeFormat: "modern",
  position: { x: 1_234_500, y: 2_345_600, unit: "emu" },
  comments: [{
    nativeId: "{11111111-1111-4111-8111-111111111111}",
    author: "Review Owner",
    person: {
      id: "{AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}",
      name: "Review Owner",
      initials: "RO",
      userId: "review.owner@example.test",
      providerId: "None",
    },
    text: "Confirm the customer evidence.",
    created: "2026-07-19T02:55:00Z",
    status: "active",
  }],
});
modernCommentThread.addReply("Evidence is attached.", {
  nativeId: "{22222222-2222-4222-8222-222222222222}",
  author: "Evidence Owner",
  person: {
    id: "{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}",
    name: "Evidence Owner",
    initials: "EO",
    userId: "evidence.owner@example.test",
    providerId: "None",
  },
  created: "2026-07-19T03:05:00Z",
  status: "active",
});
assert.equal(modernCommentDeck.verify().ok, true);
const modernCommentExport = await PresentationFile.exportPptx(modernCommentDeck);
const modernCommentZip = await JSZip.loadAsync(new Uint8Array(await modernCommentExport.arrayBuffer()));
const modernCommentPartPath = Object.keys(modernCommentZip.files).find((name) => /^ppt\/comments\/(?:modernComment|comment)\d*\.xml$/.test(name));
const modernAuthorsPartPath = Object.keys(modernCommentZip.files).find((name) => /^ppt\/(?:authors|authors\/author\d+)\.xml$/.test(name));
assert.ok(modernCommentPartPath);
assert.ok(modernAuthorsPartPath);
const modernCommentXml = await modernCommentZip.file(modernCommentPartPath).async("text");
const modernAuthorsXml = await modernCommentZip.file(modernAuthorsPartPath).async("text");
assert.match(modernCommentXml, /<p188:replyLst>/);
assert.match(modernCommentXml, /<oac:txMkLst>/);
assert.match(modernCommentXml, /<oac:txMk cp="0" len="26"/);
assert.match(modernAuthorsXml, /Review Owner/);
assert.match(modernAuthorsXml, /Evidence Owner/);

const modernCommentImported = await PresentationFile.importPptx(modernCommentExport);
assert.equal(modernCommentImported.commentFormat, "modern");
const importedModernSlide = modernCommentImported.slides.getItem(0);
const importedModernThread = importedModernSlide.comments.items[0];
assert.equal(importedModernThread.nativeFormat, "modern");
assert.equal(importedModernThread.comments.length, 2);
assert.equal(importedModernThread.comments[0].author, "Review Owner");
assert.equal(importedModernThread.comments[1].author, "Evidence Owner");
assert.equal(importedModernThread.nativeAnchor.type, "textRange");
assert.equal(importedModernThread.nativeAnchor.textLength, 26);
assert.equal(importedModernSlide.resolve(importedModernThread.targetId).kind, "textRange");

const modernUnchanged = await PresentationFile.exportPptx(modernCommentImported);
const modernUnchangedZip = await JSZip.loadAsync(new Uint8Array(await modernUnchanged.arrayBuffer()));
assert.deepEqual(
  await modernUnchangedZip.file(modernCommentPartPath).async("uint8array"),
  await modernCommentZip.file(modernCommentPartPath).async("uint8array"),
);
assert.deepEqual(
  await modernUnchangedZip.file(modernAuthorsPartPath).async("uint8array"),
  await modernCommentZip.file(modernAuthorsPartPath).async("uint8array"),
);

importedModernThread.comments[0].text = "Customer evidence confirmed.";
importedModernThread.comments[1].text = "Recorded in the decision log.";
importedModernThread.resolve();
const modernEdited = await PresentationFile.exportPptx(modernCommentImported);
const modernEditedRoundTrip = await PresentationFile.importPptx(modernEdited);
const editedModernThread = modernEditedRoundTrip.slides.getItem(0).comments.items[0];
assert.equal(editedModernThread.comments[0].text, "Customer evidence confirmed.");
assert.equal(editedModernThread.comments[1].text, "Recorded in the decision log.");
assert.equal(editedModernThread.resolved, true);
assert.equal(editedModernThread.comments[0].status, "resolved");

editedModernThread.comments[0].author = "Changed identity";
await assert.rejects(
  () => PresentationFile.exportPptx(modernEditedRoundTrip),
  (error) => error?.code === "presentation_comment_topology_changed",
);

const invalidModernCommentDeck = Presentation.create({ commentFormat: "modern" });
const invalidModernCommentSlide = invalidModernCommentDeck.slides.add();
const invalidModernTarget = invalidModernCommentSlide.shapes.add({ text: "Short" });
invalidModernCommentSlide.comments.addThread(`${invalidModernTarget.id}/text`, "Out of bounds.", {
  nativeFormat: "modern",
  nativeAnchor: { type: "textRange", cp: 3, length: 99 },
  author: "Reviewer",
  created: "2026-07-19T04:00:00Z",
  position: { x: 100, y: 100, unit: "emu" },
});
await assert.rejects(
  () => PresentationFile.exportPptx(invalidModernCommentDeck),
  (error) => error?.code === "invalid_presentation_modern_comment",
);

const missingModernCommentPositionDeck = Presentation.create({ commentFormat: "modern" });
const missingModernCommentPositionSlide = missingModernCommentPositionDeck.slides.add();
const missingModernCommentPositionTarget = missingModernCommentPositionSlide.shapes.add({ text: "Coordinate required" });
missingModernCommentPositionSlide.comments.addThread(missingModernCommentPositionTarget, "No implicit origin.", {
  nativeFormat: "modern",
  author: "Reviewer",
  created: "2026-07-19T04:10:00Z",
});
await assert.rejects(
  () => PresentationFile.exportPptx(missingModernCommentPositionDeck),
  (error) => error?.code === "invalid_presentation_modern_comment" && /explicit.*position/i.test(error.message),
);

const elementDeleteSource = Presentation.create({ slideSize: { width: 640, height: 360 } });
const elementDeleteSlide = elementDeleteSource.slides.add({ name: "Element deletion" });
elementDeleteSlide.shapes.add({ name: "keep-shape", text: "Keep", position: { left: 40, top: 40, width: 180, height: 80 } });
elementDeleteSlide.shapes.add({ name: "delete-shape", text: "Delete", position: { left: 280, top: 40, width: 180, height: 80 } });
const elementDeleteSourceFile = await PresentationFile.exportPptx(elementDeleteSource);
const elementDeleteImported = await PresentationFile.importPptx(elementDeleteSourceFile);
const importedDeleteShape = elementDeleteImported.slides.getItem(0).shapes.getItem("delete-shape");
assert.equal(importedDeleteShape.deletionCapability.sourceBound, true);
assert.equal(importedDeleteShape.deletionCapability.known, true);
assert.equal(importedDeleteShape.deletionCapability.supported, true);
assert.equal(importedDeleteShape.deletionCapability.blockedReason, "");
assert.ok(Number.isInteger(importedDeleteShape.deletionCapability.nativeId) && importedDeleteShape.deletionCapability.nativeId > 0);
assert.equal(importedDeleteShape.inspectRecord().deletionCapability.supported, true);
const deletedId = importedDeleteShape.id;
assert.equal(importedDeleteShape.delete(), importedDeleteShape);
assert.equal(elementDeleteImported.resolve(deletedId), undefined);
const elementDeletedFile = await PresentationFile.exportPptx(elementDeleteImported);
const elementDeleteRoundTrip = await PresentationFile.importPptx(elementDeletedFile);
assert.deepEqual(elementDeleteRoundTrip.slides.getItem(0).shapes.items.map((shape) => shape.name), ["keep-shape"]);

const [elementDeleteBeforeZip, elementDeleteAfterZip] = await Promise.all([
  JSZip.loadAsync(new Uint8Array(await elementDeleteSourceFile.arrayBuffer())),
  JSZip.loadAsync(new Uint8Array(await elementDeletedFile.arrayBuffer())),
]);
assert.deepEqual(Object.keys(elementDeleteAfterZip.files).sort(), Object.keys(elementDeleteBeforeZip.files).sort());
for (const name of Object.keys(elementDeleteBeforeZip.files).filter((name) => name !== "ppt/slides/slide1.xml" && !elementDeleteBeforeZip.files[name].dir)) {
  assert.deepEqual(await elementDeleteAfterZip.file(name).async("uint8array"), await elementDeleteBeforeZip.file(name).async("uint8array"), `${name} changed during shape deletion`);
}

const untypedElementDelete = await PresentationFile.importPptx(elementDeleteSourceFile);
untypedElementDelete.slides.getItem(0).shapes.items.splice(1, 1);
await assert.rejects(
  () => PresentationFile.exportPptx(untypedElementDelete),
  (error) => error?.code === "presentation_element_topology_changed",
  "raw collection mutation must not masquerade as a capability-proven element deletion",
);

const imageDeleteSource = Presentation.create({ slideSize: { width: 640, height: 360 } });
const imageDeleteSlide = imageDeleteSource.slides.add({ name: "Image deletion" });
imageDeleteSlide.images.add({ name: "keep-image", alt: "Keep", dataUrl: PNG, position: { left: 40, top: 40, width: 120, height: 120 } });
imageDeleteSlide.images.add({ name: "delete-image", alt: "Delete", dataUrl: PNG_ALT, position: { left: 240, top: 40, width: 120, height: 120 } });
const imageDeleteSourceFile = await PresentationFile.exportPptx(imageDeleteSource);
const imageDeleteImported = await PresentationFile.importPptx(imageDeleteSourceFile);
const importedDeleteImage = imageDeleteImported.slides.getItem(0).images.items.find((image) => image.name === "delete-image");
assert.equal(importedDeleteImage.deletionCapability.sourceBound, true);
assert.equal(importedDeleteImage.deletionCapability.known, true);
assert.equal(importedDeleteImage.deletionCapability.supported, true);
assert.ok(Number.isInteger(importedDeleteImage.deletionCapability.nativeId) && importedDeleteImage.deletionCapability.nativeId > 0);
assert.equal(importedDeleteImage.inspectRecord().deletionCapability.supported, true);
assert.equal(importedDeleteImage.delete(), importedDeleteImage);
const imageDeletedFile = await PresentationFile.exportPptx(imageDeleteImported);
const imageDeleteRoundTrip = await PresentationFile.importPptx(imageDeletedFile);
assert.deepEqual(imageDeleteRoundTrip.slides.getItem(0).images.items.map((image) => image.name), ["keep-image"]);

const [imageDeleteBeforeZip, imageDeleteAfterZip] = await Promise.all([
  JSZip.loadAsync(new Uint8Array(await imageDeleteSourceFile.arrayBuffer())),
  JSZip.loadAsync(new Uint8Array(await imageDeletedFile.arrayBuffer())),
]);
const imageDeleteBeforeParts = Object.keys(imageDeleteBeforeZip.files).filter((name) => !imageDeleteBeforeZip.files[name].dir);
const imageDeleteAfterParts = Object.keys(imageDeleteAfterZip.files).filter((name) => !imageDeleteAfterZip.files[name].dir);
const removedImageParts = imageDeleteBeforeParts.filter((name) => !imageDeleteAfterZip.file(name));
assert.equal(removedImageParts.filter((name) => /^ppt\/media\//.test(name)).length, 1);
const imageDeleteAllowedChanges = new Set(["[Content_Types].xml", "ppt/slides/slide1.xml", "ppt/slides/_rels/slide1.xml.rels", ...removedImageParts]);
for (const name of imageDeleteBeforeParts.filter((name) => imageDeleteAfterZip.file(name) && !imageDeleteAllowedChanges.has(name))) {
  assert.deepEqual(await imageDeleteAfterZip.file(name).async("uint8array"), await imageDeleteBeforeZip.file(name).async("uint8array"), `${name} changed during image deletion`);
}

const untypedImageDelete = await PresentationFile.importPptx(imageDeleteSourceFile);
untypedImageDelete.slides.getItem(0).images.items.splice(1, 1);
await assert.rejects(
  () => PresentationFile.exportPptx(untypedImageDelete),
  (error) => error?.code === "presentation_element_topology_changed",
  "raw image collection mutation must not masquerade as a capability-proven element deletion",
);

const structuredElementDeleteSource = Presentation.create({ slideSize: { width: 640, height: 360 } });
const structuredElementDeleteSlide = structuredElementDeleteSource.slides.add({ name: "Structured element deletion" });
structuredElementDeleteSlide.tables.add({
  name: "delete-table",
  position: { left: 40, top: 40, width: 220, height: 100 },
  values: [["Metric", "Value"], ["Pipeline", "42"]],
  styleOptions: { headerRow: true },
});
structuredElementDeleteSlide.charts.add("bar", {
  name: "delete-chart",
  title: "Pipeline",
  position: { left: 300, top: 40, width: 280, height: 180 },
  categories: ["Q1", "Q2"],
  series: [{ name: "Value", values: [42, 48] }],
});
structuredElementDeleteSlide.connectors.add({
  name: "delete-connector",
  start: { x: 60, y: 270 },
  end: { x: 560, y: 270 },
  line: { fill: "#2563eb", width: 2 },
});
structuredElementDeleteSlide.shapes.add({ name: "keep-structured-canary", text: "Keep", position: { left: 40, top: 300, width: 120, height: 40 } });
const structuredElementDeleteFile = await PresentationFile.exportPptx(structuredElementDeleteSource);
const structuredElementDeleteImported = await PresentationFile.importPptx(structuredElementDeleteFile);
const structuredImportedSlide = structuredElementDeleteImported.slides.getItem(0);
for (const [kind, element] of [
  ["table", structuredImportedSlide.tables.items[0]],
  ["chart", structuredImportedSlide.charts.items[0]],
  ["connector", structuredImportedSlide.connectors.items[0]],
]) {
  assert.equal(element.deletionCapability.sourceBound, true, `${kind} deletion must remain source-bound`);
  assert.equal(element.deletionCapability.known, true, `${kind} deletion capability must be explicit`);
  assert.equal(element.deletionCapability.supported, true, `${kind} deletion topology must be capability-proven`);
  assert.ok(Number.isInteger(element.deletionCapability.nativeId) && element.deletionCapability.nativeId > 0);
  assert.equal(element.inspectRecord().deletionCapability.supported, true);
  assert.equal(element.delete(), element);
}
const structuredElementDeletedFile = await PresentationFile.exportPptx(structuredElementDeleteImported);
const structuredElementDeleteRoundTrip = await PresentationFile.importPptx(structuredElementDeletedFile);
const structuredRoundTripSlide = structuredElementDeleteRoundTrip.slides.getItem(0);
assert.equal(structuredRoundTripSlide.tables.items.length, 0);
assert.equal(structuredRoundTripSlide.charts.items.length, 0);
assert.equal(structuredRoundTripSlide.connectors.items.length, 0);
assert.deepEqual(structuredRoundTripSlide.shapes.items.map((shape) => shape.name), ["keep-structured-canary"]);
const [structuredBeforeZip, structuredAfterZip] = await Promise.all([
  JSZip.loadAsync(new Uint8Array(await structuredElementDeleteFile.arrayBuffer())),
  JSZip.loadAsync(new Uint8Array(await structuredElementDeletedFile.arrayBuffer())),
]);
const isChartPartPath = (name) => /^ppt\/(?:slides\/)?charts\/chart\d+\.xml$/.test(name);
assert.ok(Object.keys(structuredBeforeZip.files).some(isChartPartPath));
assert.equal(Object.keys(structuredAfterZip.files).some(isChartPartPath), false, "exclusive ChartPart must be removed with its frame");

const groupDeleteSource = Presentation.create({ slideSize: { width: 640, height: 360 } });
const groupDeleteSlide = groupDeleteSource.slides.add({ name: "Group deletion" });
groupDeleteSlide.shapes.add({ name: "keep-group-canary", text: "Keep", position: { left: 20, top: 300, width: 100, height: 40 } });
const deleteGroup = groupDeleteSlide.groups.add({
  name: "delete-group",
  position: { left: 40, top: 40, width: 520, height: 220 },
  childFrame: { left: 0, top: 0, width: 520, height: 220 },
});
const deleteGroupFrom = deleteGroup.shapes.add({ name: "delete-group-from", text: "From", position: { left: 10, top: 10, width: 100, height: 50 } });
const deleteGroupTo = deleteGroup.shapes.add({ name: "delete-group-to", text: "To", position: { left: 150, top: 10, width: 100, height: 50 } });
deleteGroup.connectors.add({ name: "delete-group-connector", from: deleteGroupFrom, to: deleteGroupTo, start: { x: 110, y: 35 }, end: { x: 150, y: 35 } });
deleteGroup.images.add({ name: "delete-group-image", alt: "Delete group evidence", dataUrl: PNG, position: { left: 280, top: 10, width: 70, height: 70 } });
deleteGroup.tables.add({ name: "delete-group-table", position: { left: 10, top: 100, width: 200, height: 90 }, values: [["Gate", "State"], ["QA", "Pass"]] });
deleteGroup.charts.add("bar", { name: "delete-group-chart", position: { left: 250, top: 90, width: 240, height: 110 }, categories: ["A", "B"], series: [{ name: "Value", values: [1, 2] }] });
deleteGroup.groups.add({ name: "delete-nested-group", position: { left: 370, top: 10, width: 120, height: 60 }, childFrame: { left: 0, top: 0, width: 120, height: 60 } })
  .shapes.add({ name: "delete-nested-label", text: "Nested", position: { left: 0, top: 0, width: 120, height: 60 } });
const groupDeleteSourceFile = await PresentationFile.exportPptx(groupDeleteSource);
const groupDeleteImported = await PresentationFile.importPptx(groupDeleteSourceFile);
const importedDeleteGroup = groupDeleteImported.slides.getItem(0).groups.items.find((group) => group.name === "delete-group");
assert.equal(importedDeleteGroup.deletionCapability.sourceBound, true);
assert.equal(importedDeleteGroup.deletionCapability.known, true);
assert.equal(importedDeleteGroup.deletionCapability.supported, true, JSON.stringify(importedDeleteGroup.deletionCapability));
assert.ok(Number.isInteger(importedDeleteGroup.deletionCapability.nativeId) && importedDeleteGroup.deletionCapability.nativeId > 0);
assert.equal(importedDeleteGroup.inspectRecord().deletionCapability.supported, true);
assert.equal(importedDeleteGroup.delete(), importedDeleteGroup);
const groupDeletedFile = await PresentationFile.exportPptx(groupDeleteImported);
const groupDeleteRoundTrip = await PresentationFile.importPptx(groupDeletedFile);
assert.equal(groupDeleteRoundTrip.slides.getItem(0).groups.items.length, 0);
assert.deepEqual(groupDeleteRoundTrip.slides.getItem(0).shapes.items.map((shape) => shape.name), ["keep-group-canary"]);

const untypedGroupDelete = await PresentationFile.importPptx(groupDeleteSourceFile);
untypedGroupDelete.slides.getItem(0).groups.items.splice(0, 1);
await assert.rejects(
  () => PresentationFile.exportPptx(untypedGroupDelete),
  (error) => error?.code === "presentation_element_topology_changed",
  "raw group collection mutation must not masquerade as a capability-proven recursive deletion",
);

const commentedGroupDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
const commentedGroupSlide = commentedGroupDeck.slides.add({ name: "Comment-bound group" });
const commentedGroup = commentedGroupSlide.groups.add({ name: "commented-group", position: { left: 20, top: 20, width: 200, height: 100 } });
const commentedGroupChild = commentedGroup.shapes.add({ name: "commented-child", text: "Keep review", position: { left: 0, top: 0, width: 160, height: 60 } });
commentedGroupSlide.comments.addThread(`${commentedGroupChild.id}/text`, "Do not drop this review anchor.");
assert.throws(
  () => commentedGroup.delete(),
  (error) => error?.code === "unsupported_presentation_element_delete" && /comment targets/i.test(error.message),
  "group deletion must account for comment targets on descendants",
);

const sharedChartRelationshipZip = await JSZip.loadAsync(new Uint8Array(await structuredElementDeleteFile.arrayBuffer()));
const sharedChartSlideXml = await sharedChartRelationshipZip.file("ppt/slides/slide1.xml").async("text");
const chartFrames = sharedChartSlideXml.match(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g) || [];
const chartFrame = chartFrames.find((frame) => frame.includes('name="delete-chart"'));
assert.ok(chartFrame, "expected one native chart frame");
const duplicateChartFrame = chartFrame.replace(/id="(\d+)" name="delete-chart"/, 'id="999" name="shared-chart-relationship"');
assert.notEqual(duplicateChartFrame, chartFrame);
sharedChartRelationshipZip.file("ppt/slides/slide1.xml", sharedChartSlideXml.replace("</p:spTree>", `${duplicateChartFrame}</p:spTree>`));
const sharedChartRelationshipFile = new FileBlob(
  await sharedChartRelationshipZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }),
  { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
);
const sharedChartRelationshipImported = await PresentationFile.importPptx(sharedChartRelationshipFile);
const sharedRelationshipCharts = sharedChartRelationshipImported.slides.getItem(0).charts.items;
assert.equal(sharedRelationshipCharts.length, 2);
assert.ok(sharedRelationshipCharts.every((chart) => chart.deletionCapability.supported === false));
assert.ok(sharedRelationshipCharts.every((chart) => /referenced outside/.test(chart.deletionCapability.blockedReason)));
assert.throws(
  () => sharedRelationshipCharts[0].delete(),
  (error) => error?.code === "unsupported_presentation_element_delete",
  "a ChartPart relationship reused by another frame must fail closed",
);

const sharedImageDeleteSource = Presentation.create({ slideSize: { width: 640, height: 360 } });
sharedImageDeleteSource.slides.add({ name: "Delete shared image" }).images.add({ name: "shared-delete", dataUrl: PNG, position: { left: 40, top: 40, width: 120, height: 120 } });
sharedImageDeleteSource.slides.add({ name: "Keep shared image" }).images.add({ name: "shared-keep", dataUrl: PNG, position: { left: 40, top: 40, width: 120, height: 120 } });
const sharedImageDeleteImported = await PresentationFile.importPptx(await PresentationFile.exportPptx(sharedImageDeleteSource));
sharedImageDeleteImported.slides.getItem(0).images.items[0].delete();
const sharedImageDeleteRoundTrip = await PresentationFile.importPptx(await PresentationFile.exportPptx(sharedImageDeleteImported));
assert.equal(sharedImageDeleteRoundTrip.slides.getItem(0).images.items.length, 0);
assert.equal(sharedImageDeleteRoundTrip.slides.getItem(1).images.items[0].name, "shared-keep");

const connectedElementDeleteSource = Presentation.create({ slideSize: { width: 640, height: 360 } });
const connectedElementDeleteSlide = connectedElementDeleteSource.slides.add({ name: "Connected deletion" });
const connectedFrom = connectedElementDeleteSlide.shapes.add({ name: "connected-from", position: { left: 40, top: 40, width: 140, height: 70 } });
const connectedTo = connectedElementDeleteSlide.shapes.add({ name: "connected-to", position: { left: 360, top: 40, width: 140, height: 70 } });
connectedElementDeleteSlide.shapes.connect(connectedFrom, connectedTo);
const connectedElementDeleteImported = await PresentationFile.importPptx(await PresentationFile.exportPptx(connectedElementDeleteSource));
const blockedElementDeleteShape = connectedElementDeleteImported.slides.getItem(0).shapes.getItem("connected-from");
assert.equal(blockedElementDeleteShape.deletionCapability.supported, false);
assert.match(blockedElementDeleteShape.deletionCapability.blockedReason, /connector topology/);
assert.throws(
  () => blockedElementDeleteShape.delete(),
  (error) => error?.code === "unsupported_presentation_element_delete",
);
assert.equal(connectedElementDeleteImported.slides.getItem(0).shapes.count, 2);

console.log("presentation smoke ok");
