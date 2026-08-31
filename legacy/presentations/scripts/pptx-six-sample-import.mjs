#!/usr/bin/env node

import { createHash } from "node:crypto";
import { mkdir, readFile, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

import { PresentationFile } from "../src/presentation/index.mjs";
import { FileBlob } from "../src/shared/file-blob.mjs";
import { directPresentationChildren } from "../src/presentation/group-shapes.mjs";

const SCHEMA = "office-kit/pptx-six-sample-import-evidence/v1";
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const MAX_SOURCE_BYTES = 128 * 1024 * 1024;
const DEFAULT_ASSETS_DIR = path.resolve("tmp/reference-pptx-downloads");
const DEFAULT_OUTPUT = path.resolve("tmp/presentation-six-sample-import/evidence.v1.json");

// These are reference inputs only.  The files stay under ignored tmp/ because
// the SlidesCarnival terms do not allow unmodified redistribution.
export const SOURCES = Object.freeze([
  {
    id: "nasa-froste-thermal-design",
    fileName: "nasa-froste-thermal-design.pptx",
    sha256: "41568341c527866c9c8917229d190bca8dc99f0a137e97343b2c30d03f1db7b0",
    source: "NASA NTRS 20250007120",
  },
  {
    id: "nasa-mms-machine-learning",
    fileName: "nasa-mms-machine-learning.pptx",
    sha256: "531c82797fde09b1ebe1e868ca9cd44c3e2f675dc8f09f58b54bab6a62629723",
    source: "NASA NTRS 20250000748",
  },
  {
    id: "slidescarnival-business-infographic",
    fileName: "slidescarnival-business-infographic.pptx",
    sha256: "8db900eb9fbc5375d6b69eccffebd5ebb002f2f6641a89f19364a74e1d7e1e26",
    source: "SlidesCarnival Business Infographic",
  },
  {
    id: "slidescarnival-data-particles",
    fileName: "slidescarnival-data-particles.pptx",
    sha256: "07cd6c7e3c12335716fbfddb1ccde353c9d21959427e2639dea29eca1573464f",
    source: "SlidesCarnival Data Particles",
  },
  {
    id: "slidescarnival-minimal-business",
    fileName: "slidescarnival-minimal-business.pptx",
    sha256: "5076299de79a0c8ede47bb2f3c286a8e021ef0f7a55c991279ad03d4717ca334",
    source: "SlidesCarnival Minimal Business",
  },
  {
    id: "slidescarnival-professional-minimalist",
    fileName: "slidescarnival-professional-minimalist.pptx",
    sha256: "ed26f074e25361e4baf76d1cc2508596655b9d4f3fc2c659ecd962f76e0547ab",
    source: "SlidesCarnival Professional Minimalist",
  },
]);

export async function collectSixSampleEvidence({ assetsDir = DEFAULT_ASSETS_DIR } = {}) {
  const root = path.resolve(assetsDir);
  const results = [];
  for (const source of SOURCES) {
    const sourcePath = path.join(root, source.fileName);
    const bytes = await readPptx(sourcePath);
    const digest = sha256(bytes);
    if (digest !== source.sha256) throw new Error(`${source.id} source SHA-256 mismatch: ${digest}`);
    const packageInfo = await packageEvidence(bytes);
    const presentation = await importPresentation(bytes);
    const records = parseNdjson(presentation.inspect({ kind: "importObject", includeNested: true, maxChars: Infinity }).ndjson);
    const topLevelRecords = records.filter((record) => record.topLevel === true);
    const nestedRecords = records.filter((record) => record.topLevel === false);
    const rawObjectCount = packageInfo.rawObjectCount;
    if (presentation.slides.count !== packageInfo.slideCount) {
      throw new Error(`${source.id} imported ${presentation.slides.count} slides, expected ${packageInfo.slideCount}.`);
    }
    if (topLevelRecords.length !== rawObjectCount) {
      throw new Error(`${source.id} classified ${topLevelRecords.length} of ${rawObjectCount} visible top-level objects.`);
    }
    if (records.length !== packageInfo.visibleObjectCount) {
      throw new Error(`${source.id} classified ${records.length} of ${packageInfo.visibleObjectCount} visible objects including nested group children.`);
    }
    if (new Set(records.map((record) => record.targetId)).size !== records.length) {
      throw new Error(`${source.id} has duplicate imported object IDs.`);
    }
    const noOp = await PresentationFile.exportPptx(presentation);
    if (!Buffer.from(noOp.bytes).equals(bytes)) throw new Error(`${source.id} no-op export is not byte-identical.`);

    const profile = presentation.designProfile({ maxItems: 64, includeComponentCandidates: true });
    const placement = await verifyPlacementEdit(bytes);
    const zOrder = await verifyZOrderEdit(bytes);
    const text = await verifyTextEdit(bytes);
    const nativeText = await verifyNativeTextEdit(bytes);
    const imageReplacement = await verifyImageReplacement(bytes);
    const imageCrop = await verifyImageCropEdit(bytes);
    const nativeFill = await verifyNativeFillEdit(bytes);
    const nativeLine = await verifyNativeLineEdit(bytes);
    const nativeLineWidth = await verifyNativeLineWidthEdit(bytes);
    const nativeLineStyle = await verifyNativeLineStyleEdit(bytes);
    const nativeLineCap = await verifyNativeLineCapEdit(bytes);
    const nativeLineJoin = await verifyNativeLineJoinEdit(bytes);
    const nativeLineStartArrow = await verifyNativeLineArrowEdit(bytes, "lineStartArrow");
    const nativeLineEndArrow = await verifyNativeLineArrowEdit(bytes, "lineEndArrow");
    const nativeFontSize = await verifyNativeFontSizeEdit(bytes);
    const nativeFontFamily = await verifyNativeFontFamilyEdit(bytes);
    const nativeFontStyle = await verifyNativeFontStyleEdit(bytes);
    const nativeFontColor = await verifyNativeFontColorEdit(bytes);
    const nativeFontDecoration = await verifyNativeFontDecorationEdit(bytes);
    const nativeParagraphAlignment = await verifyNativeParagraphAlignmentEdit(bytes);
    const nativeParagraphLineSpacing = await verifyNativeParagraphLineSpacingEdit(bytes);
    const nativeParagraphBlockSpacing = await verifyNativeParagraphBlockSpacingEdit(bytes);
    const nativeParagraphMargin = await verifyNativeParagraphLayoutEdit(bytes, "paragraphMarginLeftEmu");
    const nativeParagraphIndent = await verifyNativeParagraphLayoutEdit(bytes, "paragraphIndentEmu");
    const nativeParagraphBullet = await verifyNativeParagraphBulletEdit(bytes);
    const nativeParagraphAutoNumberScheme = await verifyNativeParagraphAutoNumberEdit(bytes, "paragraphBulletAutoNumberScheme");
    const nativeParagraphAutoNumberStartAt = await verifyNativeParagraphAutoNumberEdit(bytes, "paragraphBulletAutoNumberStartAt");
    const nativeParagraphLevel = await verifyNativeParagraphLevelEdit(bytes);
    const nativeVerticalAnchor = await verifyNativeVerticalAnchorEdit(bytes);
    const nativeTextBodyInset = await verifyNativeTextBodyInsetEdit(bytes);
    const nativeTextBodyWrap = await verifyNativeTextBodyWrapEdit(bytes);
    const nativeTextBodyColumnCount = await verifyNativeTextBodyColumnCountEdit(bytes);
    const nativeTextBodyAutoFit = await verifyNativeTextBodyAutoFitEdit(bytes);
    const nativeTextBodyColumnDirection = await verifyNativeTextBodyColumnDirectionEdit(bytes);
    const nativeTextBodyVerticalText = await verifyNativeTextBodyVerticalTextEdit(bytes);
    const nativeRotation = await verifyNativeRotationEdit(bytes);
    const nativeFlip = await verifyNativeFlipEdit(bytes);
    const nativeFillOpacity = await verifyNativeFillOpacityEdit(bytes);
    const svgStyle = await verifySvgStyleEdit(bytes);
    const animatedText = await verifyAnimatedTextEdit(bytes);
    const tableCell = await verifyTableCellEdit(bytes);
    const reuse = await verifyOneSlideReuse(bytes);
    const componentReuse = await verifySourceComponentReuse(bytes);
    results.push({
      id: source.id,
      fileName: source.fileName,
      source: source.source,
      sourceSha256: digest,
      bytes: bytes.byteLength,
      slides: packageInfo.slideCount,
      visibleTopLevelObjects: rawObjectCount,
      visibleNestedObjects: packageInfo.nestedObjectCount,
      visibleObjects: packageInfo.visibleObjectCount,
      classifiedTopLevelObjects: topLevelRecords.length,
      classifiedNestedObjects: nestedRecords.length,
      classifiedObjects: records.length,
      rawRootKinds: packageInfo.rawRootKinds,
      objectKinds: counts(topLevelRecords.map((record) => record.objectKind)),
      nestedObjectKinds: counts(nestedRecords.map((record) => record.objectKind)),
      classifications: counts(topLevelRecords.map((record) => record.classification)),
      nestedClassifications: counts(nestedRecords.map((record) => record.classification)),
      nativeLeafKinds: counts(topLevelRecords.flatMap((record) => record.nativeLeafKinds || [])),
      noOpByteIdentical: true,
      placement,
      zOrder,
      text,
      nativeText,
      imageReplacement,
      imageCrop,
      nativeFill,
      nativeLine,
      nativeLineWidth,
      nativeLineStyle,
      nativeLineCap,
      nativeLineJoin,
      nativeLineStartArrow,
      nativeLineEndArrow,
      nativeFontSize,
      nativeFontFamily,
      nativeFontStyle,
      nativeFontColor,
      nativeFontDecoration,
      nativeParagraphAlignment,
      nativeParagraphLineSpacing,
      nativeParagraphBlockSpacing,
      nativeParagraphMargin,
      nativeParagraphIndent,
      nativeParagraphBullet,
      nativeParagraphAutoNumberScheme,
      nativeParagraphAutoNumberStartAt,
      nativeParagraphLevel,
      nativeVerticalAnchor,
      nativeTextBodyInset,
      nativeTextBodyWrap,
      nativeTextBodyColumnCount,
      nativeTextBodyAutoFit,
      nativeTextBodyColumnDirection,
      nativeTextBodyVerticalText,
      nativeRotation,
      nativeFlip,
      nativeFillOpacity,
      svgStyle,
      animatedText,
      tableCell,
      sourceSlideReuse: reuse,
      sourceComponentReuse: componentReuse,
      nativeLeafCount: topLevelRecords.reduce((sum, record) => sum + Number(record.nativeLeafCount || 0), 0),
      nativeNestedLeafCount: nestedRecords.reduce((sum, record) => sum + Number(record.nativeLeafCount || 0), 0),
      nativeTextLeafCount: topLevelRecords
        .filter((record) => (record.nativeLeafKinds || []).includes("nativeText"))
        .reduce((sum, record) => sum + Number(record.nativeLeafCount || 0), 0),
      designProfile: profileSummary(profile),
    });
  }
  return {
    schema: SCHEMA,
    sourcePolicy: "ignored reference inputs; do not redistribute unmodified SlidesCarnival files",
    totals: {
      sources: results.length,
      slides: results.reduce((sum, result) => sum + result.slides, 0),
      visibleTopLevelObjects: results.reduce((sum, result) => sum + result.visibleTopLevelObjects, 0),
      visibleNestedObjects: results.reduce((sum, result) => sum + result.visibleNestedObjects, 0),
      visibleObjects: results.reduce((sum, result) => sum + result.visibleObjects, 0),
      classifiedObjects: results.reduce((sum, result) => sum + result.classifiedObjects, 0),
      noOpByteIdentical: results.every((result) => result.noOpByteIdentical),
      placementEdits: results.filter((result) => result.placement.status === "passed").length,
      zOrderEdits: results.filter((result) => result.zOrder.status === "passed").length,
      textEdits: results.filter((result) => result.text.status === "passed").length,
      nativeTextEdits: results.filter((result) => result.nativeText.status === "passed").length,
      imageReplacements: results.filter((result) => result.imageReplacement.status === "passed").length,
      imageCropEdits: results.filter((result) => result.imageCrop.status === "passed").length,
      nativeFillEdits: results.filter((result) => result.nativeFill.status === "passed").length,
      nativeLineEdits: results.filter((result) => result.nativeLine.status === "passed").length,
      nativeLineWidthEdits: results.filter((result) => result.nativeLineWidth.status === "passed").length,
      nativeLineStyleEdits: results.filter((result) => result.nativeLineStyle.status === "passed").length,
      nativeLineCapEdits: results.filter((result) => result.nativeLineCap.status === "passed").length,
      nativeLineJoinEdits: results.filter((result) => result.nativeLineJoin.status === "passed").length,
      nativeLineStartArrowEdits: results.filter((result) => result.nativeLineStartArrow.status === "passed").length,
      nativeLineEndArrowEdits: results.filter((result) => result.nativeLineEndArrow.status === "passed").length,
      nativeFontSizeEdits: results.filter((result) => result.nativeFontSize.status === "passed").length,
      nativeFontFamilyEdits: results.filter((result) => result.nativeFontFamily.status === "passed").length,
      nativeFontStyleEdits: results.filter((result) => result.nativeFontStyle.status === "passed").length,
      nativeFontColorEdits: results.filter((result) => result.nativeFontColor.status === "passed").length,
      nativeFontDecorationEdits: results.filter((result) => result.nativeFontDecoration.status === "passed").length,
      nativeParagraphAlignmentEdits: results.filter((result) => result.nativeParagraphAlignment.status === "passed").length,
      nativeParagraphLineSpacingEdits: results.filter((result) => result.nativeParagraphLineSpacing.status === "passed").length,
      nativeParagraphBlockSpacingEdits: results.filter((result) => result.nativeParagraphBlockSpacing.status === "passed").length,
      nativeParagraphMarginEdits: results.filter((result) => result.nativeParagraphMargin.status === "passed").length,
      nativeParagraphIndentEdits: results.filter((result) => result.nativeParagraphIndent.status === "passed").length,
      nativeParagraphBulletEdits: results.filter((result) => result.nativeParagraphBullet.status === "passed").length,
      nativeParagraphAutoNumberSchemeEdits: results.filter((result) => result.nativeParagraphAutoNumberScheme.status === "passed").length,
      nativeParagraphAutoNumberStartAtEdits: results.filter((result) => result.nativeParagraphAutoNumberStartAt.status === "passed").length,
      nativeParagraphLevelEdits: results.filter((result) => result.nativeParagraphLevel.status === "passed").length,
      nativeVerticalAnchorEdits: results.filter((result) => result.nativeVerticalAnchor.status === "passed").length,
      nativeTextBodyInsetEdits: results.filter((result) => result.nativeTextBodyInset.status === "passed").length,
      nativeTextBodyWrapEdits: results.filter((result) => result.nativeTextBodyWrap.status === "passed").length,
      nativeTextBodyColumnCountEdits: results.filter((result) => result.nativeTextBodyColumnCount.status === "passed").length,
      nativeTextBodyAutoFitEdits: results.filter((result) => result.nativeTextBodyAutoFit.status === "passed").length,
      nativeTextBodyColumnDirectionEdits: results.filter((result) => result.nativeTextBodyColumnDirection.status === "passed").length,
      nativeTextBodyVerticalTextEdits: results.filter((result) => result.nativeTextBodyVerticalText.status === "passed").length,
      nativeRotationEdits: results.filter((result) => result.nativeRotation.status === "passed").length,
      nativeFlipEdits: results.filter((result) => result.nativeFlip.status === "passed").length,
      nativeFillOpacityEdits: results.filter((result) => result.nativeFillOpacity.status === "passed").length,
      svgStyleEdits: results.filter((result) => result.svgStyle.status === "passed").length,
      animatedTextEdits: results.filter((result) => result.animatedText.status === "passed").length,
      tableCellEdits: results.filter((result) => result.tableCell.status === "passed").length,
      sourceSlideReuse: results.filter((result) => result.sourceSlideReuse.status === "passed").length,
      sourceComponentReuse: results.filter((result) => result.sourceComponentReuse.status === "passed").length,
    },
    sources: results,
  };
}

async function verifyTextEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const target = firstTextRun(presentation);
  if (!target) return { status: "blocked", reason: "no safe text run was discovered" };
  const needle = target.run.text.trim().split(/\s+/u)[0];
  target.shape.text.replace(needle, `${needle} OfficeKit`);
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const value = reopened.resolve(target.shape.id)?.text?.value || "";
  if (!value.includes(`${needle} OfficeKit`)) throw new Error(`Text edit did not survive re-import for ${target.shape.id}.`);
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Text edit changed unexpected parts for ${target.shape.id}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.shape.id, changedParts };
}

async function verifyNativeTextEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "nativeText");
  if (!target) return { status: "blocked", reason: "no bounded native text leaf was discovered" };
  const value = `${target.value} OfficeKit`;
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const object = reopened.resolve(target.targetId);
  const leaf = object?.nativeTextLeaves?.find((candidate) => candidate.textLeafIndex === target.textLeafIndex);
  if (leaf?.text !== value) throw new Error(`Native text edit did not survive re-import for ${target.targetId}.`);
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native text edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, textLeafIndex: target.textLeafIndex, changedParts };
}

async function verifyImageReplacement(bytes) {
  const presentation = await importPresentation(bytes);
  const pair = firstReplaceableImagePair(presentation);
  if (!pair) return { status: "blocked", reason: "no two distinct same-format images were discovered" };
  const { target, replacement } = pair;
  const contentType = String(target.contentType || "").toLowerCase();
  if (!contentType || contentType !== String(replacement.contentType || "").toLowerCase()) {
    throw new Error(`Imported image ${target.id} did not expose a stable matching content type.`);
  }
  const replacementBytes = dataUrlBytes(replacement.dataUrl);
  const previousCrop = target.crop ? { ...target.crop } : undefined;
  target.replace({ blob: new FileBlob(replacementBytes, { type: contentType }) });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = reopened.resolve(target.id);
  if (!rebound || rebound.contentType !== contentType || rebound.dataUrl !== replacement.dataUrl) {
    throw new Error(`Image replacement did not survive re-import for ${target.id}.`);
  }
  if (JSON.stringify(rebound.crop) !== JSON.stringify(previousCrop)) {
    throw new Error(`Image replacement changed the crop for ${target.id}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedSlide = `ppt/slides/slide${target.slide.index + 1}.xml`;
  const expectedRels = `ppt/slides/_rels/slide${target.slide.index + 1}.xml.rels`;
  const addedMedia = changedParts.filter((part) => /^ppt\/media\/office-kit-[0-9a-f]+\.(?:png|jpe?g|gif|svg)$/u.test(part));
  if (changedParts.length !== 3 || !changedParts.includes(expectedSlide) || !changedParts.includes(expectedRels) || addedMedia.length !== 1) {
    throw new Error(`Image replacement changed unexpected parts for ${target.id}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.id, replacementId: replacement.id, contentType, changedParts };
}

async function verifyImageCropEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const target = presentation.slides.items
    .flatMap((slide) => slide.images.items.map((image) => ({ slide, image })))
    .find(({ image }) => image.contentType && image.dataUrl);
  if (!target) return { status: "blocked", reason: "no embedded image with a bounded crop surface was discovered" };
  const before = target.image.crop ? { ...target.image.crop } : undefined;
  const crop = { left: 0.06, top: 0.03, right: 0.02, bottom: 0.05 };
  target.image.crop = crop;
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = reopened.resolve(target.image.id);
  if (!rebound || JSON.stringify(rebound.crop) !== JSON.stringify(crop)) {
    throw new Error(`Image crop edit did not survive re-import for ${target.image.id}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide.index + 1}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Image crop edit changed unexpected parts for ${target.image.id}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.image.id, before, crop, changedParts };
}

async function verifyNativeFillEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "fillRgb");
  if (!target) return { status: "blocked", reason: "no bounded fill color leaf was discovered" };
  const value = target.value.toLowerCase() === "#aabbcc" ? "#C3B2A1" : "#AABBCC";
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === "fillRgb");
  if (!rebound || rebound.value.toLowerCase() !== value.toLowerCase()) {
    throw new Error(`Native fill edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native fill edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, oldValue: target.value, value: value.toLowerCase(), changedParts };
}

async function verifyNativeLineEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "lineRgb" || record.leafKind === "lineScheme");
  if (!target) return { status: "blocked", reason: "no bounded imported connector line-color leaf was discovered" };
  const value = target.leafKind === "lineScheme"
    ? (target.value.toLowerCase() === "accent1" ? "accent2" : "accent1")
    : (target.value.toLowerCase() === "#aabbcc" ? "#C3B2A1" : "#AABBCC");
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === target.leafKind);
  if (!rebound || rebound.value.toLowerCase() !== value.toLowerCase()) {
    throw new Error(`Native connector line edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native connector line edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, oldValue: target.value, value: value.toLowerCase(), changedParts };
}

async function verifyNativeLineWidthEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "lineWidthEmu");
  if (!target) return { status: "blocked", reason: "no bounded imported line-width leaf was discovered" };
  const oldValue = Number(target.value);
  const value = oldValue + 9525 <= 20_116_800 ? oldValue + 9525 : Math.max(0, oldValue - 9525);
  if (!Number.isSafeInteger(oldValue) || value === oldValue) return { status: "blocked", reason: "discovered line width is outside the safe edit range" };
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value: String(value) });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === "lineWidthEmu");
  if (!rebound || Number(rebound.value) !== value) {
    throw new Error(`Native line-width edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native line-width edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, oldValue, value, changedParts };
}

async function verifyNativeLineStyleEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "lineStyle");
  if (!target) return { status: "blocked", reason: "no bounded imported preset dash leaf was discovered" };
  const value = target.value === "solid" ? "dashed" : "solid";
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === "lineStyle");
  if (!rebound || rebound.value !== value) {
    throw new Error(`Native line-style edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native line-style edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, oldValue: target.value, value, changedParts };
}

async function verifyNativeLineCapEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "lineCap");
  if (!target) return { status: "blocked", reason: "no bounded imported line-cap leaf was discovered" };
  const value = target.value === "flat" ? "round" : "flat";
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === "lineCap");
  if (!rebound || rebound.value !== value) {
    throw new Error(`Native line-cap edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native line-cap edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, oldValue: target.value, value, changedParts };
}

async function verifyNativeLineJoinEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "lineJoin");
  if (!target) return { status: "blocked", reason: "no bounded imported line-join leaf was discovered" };
  const value = target.value === "round" ? "bevel" : "round";
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === "lineJoin");
  if (!rebound || rebound.value !== value) {
    throw new Error(`Native line-join edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native line-join edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, oldValue: target.value, value, changedParts };
}

async function verifyNativeLineArrowEdit(bytes, leafKind) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === leafKind);
  if (!target) return { status: "blocked", reason: `no bounded imported ${leafKind} leaf was discovered` };
  const value = target.value === "none" ? "triangle" : "none";
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === leafKind);
  if (!rebound || rebound.value !== value) {
    throw new Error(`Native ${leafKind} edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native ${leafKind} edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, oldValue: target.value, value, changedParts };
}

async function verifyNativeFontSizeEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "fontSizePoints");
  if (!target) return { status: "blocked", reason: "no bounded explicit text run font-size leaf was discovered" };
  const oldValue = Number(target.value);
  const value = oldValue + 1 <= 768 ? Math.round((oldValue + 1) * 100) / 100 : Math.round((oldValue - 1) * 100) / 100;
  if (!Number.isFinite(oldValue) || value <= 0 || value === oldValue) return { status: "blocked", reason: "discovered font size is outside the safe edit range" };
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === "fontSizePoints");
  if (!rebound || Math.abs(Number(rebound.value) - value) > 0.001) {
    throw new Error(`Native font-size edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native font-size edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, textLeafIndex: target.textLeafIndex, oldValue, value, changedParts };
}

async function verifyNativeFontFamilyEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "fontFamily" || record.leafKind === "fontFamilyEastAsia");
  if (!target) return { status: "blocked", reason: "no bounded explicit text run font-family leaf was discovered" };
  const value = target.value === "OfficeKit Sans" ? "Arial" : "OfficeKit Sans";
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === target.leafKind);
  if (!rebound || rebound.value !== value) {
    throw new Error(`Native font-family edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native font-family edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind: target.leafKind, textLeafIndex: target.textLeafIndex, oldValue: target.value, value, changedParts };
}

async function verifyNativeFontStyleEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "fontBold" || record.leafKind === "fontItalic");
  if (!target) return { status: "blocked", reason: "no bounded explicit text run bold/italic leaf was discovered" };
  const value = !target.value;
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === target.leafKind);
  if (!rebound || rebound.value !== value) {
    throw new Error(`Native font-style edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native font-style edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind: target.leafKind, textLeafIndex: target.textLeafIndex, oldValue: target.value, value, changedParts };
}

async function verifyNativeFontColorEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "fontColorRgb");
  if (!target) return { status: "blocked", reason: "no bounded bare RGB text-run color leaf was discovered" };
  const value = target.value.toLowerCase() === "#aabbcc" ? "#C3B2A1" : "#AABBCC";
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === "fontColorRgb");
  if (!rebound || rebound.value.toLowerCase() !== value.toLowerCase()) {
    throw new Error(`Native font-color edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native font-color edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind: target.leafKind, textLeafIndex: target.textLeafIndex, oldValue: target.value, value, changedParts };
}

async function verifyNativeFontDecorationEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "fontUnderline" || record.leafKind === "fontStrike");
  if (!target) return { status: "blocked", reason: "no bounded direct run underline or strike leaf was discovered" };
  const value = target.leafKind === "fontUnderline"
    ? (target.value === "sng" ? "dbl" : target.value === "none" ? "sng" : "none")
    : (target.value === "noStrike" ? "sngStrike" : target.value === "sngStrike" ? "dblStrike" : "noStrike");
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === target.leafKind);
  if (!rebound || rebound.value !== value) {
    throw new Error(`Native font-decoration edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native font-decoration edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind: target.leafKind, textLeafIndex: target.textLeafIndex, oldValue: target.value, value, changedParts };
}

async function verifyNativeParagraphAlignmentEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "paragraphAlignment");
  if (!target) return { status: "blocked", reason: "no bounded direct paragraph-alignment leaf was discovered" };
  const value = target.value === "left" ? "center" : "left";
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === "paragraphAlignment" && Number(record.nativeLeafIndex) === Number(target.nativeLeafIndex));
  if (!rebound || rebound.value !== value) {
    throw new Error(`Native paragraph-alignment edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native paragraph-alignment edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, nativeLeafIndex: target.nativeLeafIndex, oldValue: target.value, value, changedParts };
}

async function verifyNativeParagraphLineSpacingEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "paragraphLineSpacingPoints" || record.leafKind === "paragraphLineSpacingMultiplier");
  if (!target) return { status: "blocked", reason: "no bounded direct paragraph line-spacing leaf was discovered" };
  const oldValue = Number(target.value);
  const maximum = target.leafKind === "paragraphLineSpacingPoints" ? 1584 : 132;
  const step = 0.01;
  const value = oldValue + step <= maximum ? Number((oldValue + step).toFixed(5)) : Number((oldValue - step).toFixed(5));
  if (!Number.isFinite(oldValue) || value <= 0 || value === oldValue) return { status: "blocked", reason: "discovered paragraph line spacing is outside the safe edit range" };
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === target.leafKind && Number(record.nativeLeafIndex) === Number(target.nativeLeafIndex));
  if (!rebound || Math.abs(Number(rebound.value) - value) > 0.00001) {
    throw new Error(`Native paragraph line-spacing edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native paragraph line-spacing edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind: target.leafKind, nativeLeafIndex: target.nativeLeafIndex, oldValue, value, changedParts };
}

async function verifyNativeParagraphBlockSpacingEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => [
    "paragraphSpaceBeforePoints",
    "paragraphSpaceBeforeMultiplier",
    "paragraphSpaceAfterPoints",
    "paragraphSpaceAfterMultiplier",
  ].includes(record.leafKind));
  if (!target) return { status: "blocked", reason: "no bounded direct paragraph before/after-spacing leaf was discovered" };
  const oldValue = Number(target.value);
  const maximum = target.leafKind.endsWith("Points") ? 1584 : 132;
  const step = 0.01;
  const value = oldValue + step <= maximum
    ? Number((oldValue + step).toFixed(5))
    : Number((oldValue - step).toFixed(5));
  if (!Number.isFinite(oldValue) || value < 0 || value === oldValue) {
    return { status: "blocked", reason: "discovered paragraph before/after spacing is outside the safe edit range" };
  }
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === target.leafKind && Number(record.nativeLeafIndex) === Number(target.nativeLeafIndex));
  if (!rebound || Math.abs(Number(rebound.value) - value) > 0.00001) {
    throw new Error(`Native paragraph before/after-spacing edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native paragraph before/after-spacing edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind: target.leafKind, nativeLeafIndex: target.nativeLeafIndex, oldValue, value, changedParts };
}

async function verifyNativeParagraphLayoutEdit(bytes, leafKind) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === leafKind);
  if (!target) return { status: "blocked", reason: `no bounded direct ${leafKind} leaf was discovered` };
  const oldValue = Number(target.value);
  const maximum = 51_206_400;
  const value = oldValue < maximum ? oldValue + 1 : oldValue - 1;
  if (!Number.isSafeInteger(oldValue) || !Number.isSafeInteger(value) || value === oldValue || value < -maximum || value > maximum || (leafKind === "paragraphMarginLeftEmu" && value < 0)) {
    return { status: "blocked", reason: `discovered ${leafKind} is outside the safe edit range` };
  }
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === leafKind && Number(record.nativeLeafIndex) === Number(target.nativeLeafIndex));
  if (!rebound || Number(rebound.value) !== value) {
    throw new Error(`Native ${leafKind} edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native ${leafKind} edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind, nativeLeafIndex: target.nativeLeafIndex, oldValue, value, changedParts };
}

async function verifyNativeParagraphBulletEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "paragraphBulletCharacter");
  if (!target) return { status: "blocked", reason: "no bounded direct character-bullet leaf was discovered" };
  const oldValue = String(target.value);
  const value = oldValue === "•" ? "◦" : "•";
  if ([...oldValue].length !== 1 || [...value].length !== 1 || value === oldValue) {
    return { status: "blocked", reason: "discovered character bullet has no safe alternate scalar" };
  }
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === "paragraphBulletCharacter" && Number(record.nativeLeafIndex) === Number(target.nativeLeafIndex));
  if (!rebound || rebound.value !== value) {
    throw new Error(`Native character-bullet edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native character-bullet edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind: target.leafKind, nativeLeafIndex: target.nativeLeafIndex, oldValue, value, changedParts };
}

async function verifyNativeParagraphAutoNumberEdit(bytes, leafKind) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === leafKind);
  if (!target) return { status: "blocked", reason: `no bounded direct ${leafKind} leaf was discovered` };
  const oldValue = leafKind === "paragraphBulletAutoNumberScheme" ? String(target.value) : Number(target.value);
  const value = leafKind === "paragraphBulletAutoNumberScheme"
    ? (oldValue === "arabicPeriod" ? "romanLcPeriod" : "arabicPeriod")
    : (oldValue < 32767 ? oldValue + 1 : oldValue - 1);
  if (value === oldValue) return { status: "blocked", reason: `discovered ${leafKind} has no safe alternate value` };
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === leafKind && Number(record.nativeLeafIndex) === Number(target.nativeLeafIndex));
  if (!rebound || (leafKind === "paragraphBulletAutoNumberScheme" ? rebound.value !== value : Number(rebound.value) !== value)) {
    throw new Error(`Native ${leafKind} edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native ${leafKind} edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind, nativeLeafIndex: target.nativeLeafIndex, oldValue, value, changedParts };
}

async function verifyNativeParagraphLevelEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "paragraphLevel");
  if (!target) return { status: "blocked", reason: "no explicit non-zero paragraph-level leaf was discovered" };
  const oldValue = Number(target.value);
  const value = oldValue < 8 ? oldValue + 1 : oldValue - 1;
  if (!Number.isInteger(oldValue) || oldValue < 1 || oldValue > 8 || value < 1 || value > 8 || value === oldValue) {
    return { status: "blocked", reason: "discovered paragraph level has no safe successor value" };
  }
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === "paragraphLevel" && Number(record.nativeLeafIndex) === Number(target.nativeLeafIndex));
  if (!rebound || Number(rebound.value) !== value) {
    throw new Error(`Native paragraph-level edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native paragraph-level edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind: target.leafKind, nativeLeafIndex: target.nativeLeafIndex, oldValue, value, changedParts };
}

async function verifyNativeVerticalAnchorEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "verticalAnchor");
  if (!target) return { status: "blocked", reason: "no bounded direct text-body vertical-anchor leaf was discovered" };
  const value = target.value === "top" ? "center" : "top";
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === "verticalAnchor" && Number(record.nativeLeafIndex) === Number(target.nativeLeafIndex));
  if (!rebound || rebound.value !== value) {
    throw new Error(`Native vertical-anchor edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native vertical-anchor edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, nativeLeafIndex: target.nativeLeafIndex, oldValue: target.value, value, changedParts };
}

async function verifyNativeTextBodyInsetEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "textBodyInsetLeftEmu");
  if (!target) return { status: "blocked", reason: "no bounded direct text-body inset leaf was discovered" };
  const value = Number(target.value) + 1;
  if (!Number.isSafeInteger(value) || value < 0) return { status: "blocked", reason: "text-body inset leaf has no safe successor value" };
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === target.leafKind && Number(record.nativeLeafIndex) === Number(target.nativeLeafIndex));
  if (!rebound || Number(rebound.value) !== value) {
    throw new Error(`Native text-body inset edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native text-body inset edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind: target.leafKind, nativeLeafIndex: target.nativeLeafIndex, oldValue: target.value, value, changedParts };
}

async function verifyNativeTextBodyWrapEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "textBodyWrap");
  if (!target) return { status: "blocked", reason: "no bounded direct text-body wrap leaf was discovered" };
  const value = target.value === "square" ? "none" : "square";
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === target.leafKind && Number(record.nativeLeafIndex) === Number(target.nativeLeafIndex));
  if (!rebound || rebound.value !== value) {
    throw new Error(`Native text-body wrap edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native text-body wrap edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind: target.leafKind, nativeLeafIndex: target.nativeLeafIndex, oldValue: target.value, value, changedParts };
}

async function verifyNativeTextBodyColumnCountEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "textBodyColumnCount");
  if (!target) return { status: "blocked", reason: "no bounded direct text-body column-count leaf was discovered" };
  const oldValue = Number(target.value);
  const value = oldValue === 1 ? 2 : 1;
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === target.leafKind && Number(record.nativeLeafIndex) === Number(target.nativeLeafIndex));
  if (!rebound || Number(rebound.value) !== value) {
    throw new Error(`Native text-body column-count edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native text-body column-count edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind: target.leafKind, nativeLeafIndex: target.nativeLeafIndex, oldValue, value, changedParts };
}

async function verifyNativeTextBodyAutoFitEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "textBodyAutoFit");
  if (!target) return { status: "blocked", reason: "no bounded direct text-body AutoFit leaf was discovered" };
  const value = target.value === "none" ? "resizeShape" : "none";
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === target.leafKind && Number(record.nativeLeafIndex) === Number(target.nativeLeafIndex));
  if (!rebound || rebound.value !== value) {
    throw new Error(`Native text-body AutoFit edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native text-body AutoFit edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind: target.leafKind, nativeLeafIndex: target.nativeLeafIndex, oldValue: target.value, value, changedParts };
}

async function verifyNativeTextBodyColumnDirectionEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "textBodyColumnDirection");
  if (!target) return { status: "blocked", reason: "no bounded direct text-body column-direction leaf was discovered" };
  const value = !Boolean(target.value);
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === target.leafKind && Number(record.nativeLeafIndex) === Number(target.nativeLeafIndex));
  if (!rebound || rebound.value !== value) {
    throw new Error(`Native text-body column-direction edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native text-body column-direction edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind: target.leafKind, nativeLeafIndex: target.nativeLeafIndex, oldValue: target.value, value, changedParts };
}

async function verifyNativeTextBodyVerticalTextEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "textBodyVerticalText");
  if (!target) return { status: "blocked", reason: "no bounded direct text-body vertical-text leaf was discovered" };
  const value = target.value === "horizontal" ? "vertical" : "horizontal";
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === target.leafKind && Number(record.nativeLeafIndex) === Number(target.nativeLeafIndex));
  if (!rebound || rebound.value !== value) {
    throw new Error(`Native text-body vertical-text edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native text-body vertical-text edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind: target.leafKind, nativeLeafIndex: target.nativeLeafIndex, oldValue: target.value, value, changedParts };
}

async function verifyNativeRotationEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "rotationDegrees");
  if (!target) return { status: "blocked", reason: "no bounded direct a:xfrm rotation leaf was discovered" };
  const oldValue = Number(target.value);
  const value = oldValue === 0 ? 1 : 0;
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === target.leafKind && Number(record.nativeLeafIndex) === Number(target.nativeLeafIndex));
  if (!rebound || Number(rebound.value) !== value) {
    throw new Error(`Native rotation edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native rotation edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind: target.leafKind, nativeLeafIndex: target.nativeLeafIndex, oldValue, value, changedParts };
}

async function verifyNativeFlipEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "flipHorizontal" || record.leafKind === "flipVertical");
  if (!target) return { status: "blocked", reason: "no bounded direct a:xfrm flip leaf was discovered" };
  const value = !Boolean(target.value);
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === target.leafKind && Number(record.nativeLeafIndex) === Number(target.nativeLeafIndex));
  if (!rebound || rebound.value !== value) {
    throw new Error(`Native ${target.leafKind} edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native ${target.leafKind} edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind: target.leafKind, nativeLeafIndex: target.nativeLeafIndex, oldValue: target.value, value, changedParts };
}

async function verifyNativeFillOpacityEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson);
  const target = records.find((record) => record.leafKind === "fillOpacityThousandthPercent");
  if (!target) return { status: "blocked", reason: "no bounded direct solid-fill opacity leaf was discovered" };
  const oldValue = Number(target.value);
  const value = oldValue > 0.5 ? 0.35 : 0.65;
  presentation.editNativeLeaf(target.targetId, target.leafId, { expectedHash: target.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = parseNdjson(reopened.inspect({ kind: "nativeLeaf", maxChars: Infinity }).ndjson)
    .find((record) => record.targetId === target.targetId && record.leafKind === target.leafKind && Number(record.nativeLeafIndex) === Number(target.nativeLeafIndex));
  if (!rebound || Math.abs(Number(rebound.value) - value) > 1e-9) {
    throw new Error(`Native fill opacity edit did not survive re-import for ${target.targetId}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Native fill opacity edit changed unexpected parts for ${target.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.targetId, leafKind: target.leafKind, nativeLeafIndex: target.nativeLeafIndex, oldValue, value, changedParts };
}

async function verifySvgStyleEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const records = parseNdjson(presentation.inspect({ kind: "image", maxChars: Infinity }).ndjson);
  let target;
  let leaf;
  let image;
  for (const record of records) {
    const candidate = presentation.resolve(record.id);
    const leaves = candidate?.getSvgEditLeaves?.() || [];
    const styleLeaf = leaves.find((item) => item.leafKind === "svgFillRgb" || item.leafKind === "svgStrokeRgb");
    if (styleLeaf) {
      target = record;
      leaf = styleLeaf;
      image = candidate;
      break;
    }
  }
  if (!target || !leaf || !image) return { status: "blocked", reason: "no safe SVG style leaf was discovered" };
  const value = leaf.value.toLowerCase() === "#aabbcc" ? "#C3B2A1" : "#AABBCC";
  image.editSvgLeaf(leaf.id, { expectedHash: leaf.expectedHash, value });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = reopened.resolve(target.id);
  const reboundLeaf = rebound?.getSvgEditLeaves?.().find((item) => item.leafKind === leaf.leafKind && item.value.toLowerCase() === value.toLowerCase());
  if (!reboundLeaf) throw new Error(`SVG style edit did not survive re-import for ${target.id}.`);
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedSlide = `ppt/slides/slide${target.slide}.xml`;
  const expectedRels = `ppt/slides/_rels/slide${target.slide}.xml.rels`;
  const addedMedia = changedParts.filter((part) => /^ppt\/media\/office-kit-[0-9a-f]+\.svg$/u.test(part));
  if (changedParts.length !== 3 || !changedParts.includes(expectedSlide) || !changedParts.includes(expectedRels) || addedMedia.length !== 1) {
    throw new Error(`SVG style edit changed unexpected parts for ${target.id}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.id, leafKind: leaf.leafKind, oldValue: leaf.value, value: value.toLowerCase(), changedParts };
}

async function verifyAnimatedTextEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const animations = parseNdjson(presentation.inspect({ kind: "animation", maxChars: Infinity }).ndjson);
  const target = animations.map((animation) => ({ animation, object: presentation.resolve(animation.targetId) }))
    .find(({ object }) => object?.text?.paragraphs?.some((paragraph) => paragraph.runs?.some((run) => typeof run.text === "string" && run.text.trim().length > 0)));
  if (!target) return { status: "blocked", reason: "no animated text target with a writable run was discovered" };
  const run = target.object.text.paragraphs.flatMap((paragraph) => paragraph.runs || [])
    .find((candidate) => typeof candidate.text === "string" && candidate.text.trim().length > 0);
  const replacement = `${run.text} OfficeKit`;
  const beforeAnimations = normalizeAnimationEvidence(animations);
  target.object.text.replace(run.text, replacement);
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = reopened.resolve(target.animation.targetId);
  if (!rebound?.text?.value?.includes(replacement)) throw new Error(`Animated text edit did not survive re-import for ${target.animation.targetId}.`);
  const afterAnimations = normalizeAnimationEvidence(parseNdjson(reopened.inspect({ kind: "animation", maxChars: Infinity }).ndjson));
  if (JSON.stringify(beforeAnimations) !== JSON.stringify(afterAnimations)) throw new Error(`Animation graph changed during text edit for ${target.animation.targetId}.`);
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.animation.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Animated text edit changed unexpected parts for ${target.animation.targetId}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.animation.targetId, animationId: target.animation.id, animationCount: animations.length, changedParts };
}

function normalizeAnimationEvidence(records) {
  return records.map((record) => {
    const normalized = { ...record };
    if (normalized.capability) {
      normalized.capability = { ...normalized.capability };
      delete normalized.capability.sourceRevisionSha256;
    }
    return normalized;
  });
}

async function verifyTableCellEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const target = firstTableCell(presentation);
  if (!target) return { status: "blocked", reason: "no table with a writable text cell was discovered" };
  const value = `${target.value} OfficeKit`;
  target.table.getCell(target.row, target.column).value = value;
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = reopened.resolve(target.table.id);
  if (!rebound || rebound.values[target.row]?.[target.column] !== value) {
    throw new Error(`Table cell edit did not survive re-import for ${target.table.id}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.table.slide.index + 1}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Table cell edit changed unexpected parts for ${target.table.id}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.table.id, row: target.row, column: target.column, changedParts };
}

async function verifyPlacementEdit(bytes) {
  const presentation = await importPresentation(bytes);
  const target = firstPlacementObject(presentation);
  if (!target) return { status: "blocked", reason: "no bounded placement capability was discovered" };
  const before = { ...target.object.position };
  target.object.setPosition({ left: before.left + 3, top: before.top + 3 });
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const rebound = reopened.resolve(target.object.id);
  if (!rebound || Math.abs(rebound.position.left - before.left - 3) > 0.01 || Math.abs(rebound.position.top - before.top - 3) > 0.01) {
    throw new Error(`Placement edit did not survive re-import for ${target.object.id}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Placement edit changed unexpected parts for ${target.object.id}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", targetId: target.object.id, nativeKind: target.object.nativeKind, changedParts };
}

async function verifyZOrderEdit(bytes) {
  const presentation = await importPresentation(bytes);
  let target;
  for (const slide of presentation.slides.items) {
    const items = slide.elements.items;
    for (let index = 0; index + 1 < items.length; index += 1) {
      if (items[index].zOrderCapability?.editable === true && items[index + 1].zOrderCapability?.editable === true) {
        target = { slide, first: items[index], second: items[index + 1] };
        break;
      }
    }
    if (target) break;
  }
  if (!target) {
    for (const slide of presentation.slides.items) {
      for (const group of slide.groups?.items || []) {
        const items = group.children || [];
        for (let index = 0; index + 1 < items.length; index += 1) {
          if (items[index].zOrderCapability?.editable === true && items[index + 1].zOrderCapability?.editable === true) {
            target = { slide, group, first: items[index], second: items[index + 1] };
            break;
          }
        }
        if (target) break;
      }
      if (target) break;
    }
  }
  if (!target) return { status: "blocked", reason: "no adjacent direct or grouped elements exposed a safe z-order capability" };
  const identity = (element) => target.group
    ? `${element.kind || ""}\0${element.name || ""}\0${element.text?.value || ""}`
    : `${element.nativeId ?? ""}\0${element.name || ""}`;
  const before = (target.group ? target.group.children : target.slide.elements.items).map(identity);
  target.first.moveAfter(target.second);
  const expected = [...before];
  const firstIndex = expected.indexOf(identity(target.first));
  const secondIndex = expected.indexOf(identity(target.second));
  expected.splice(firstIndex, 1);
  expected.splice(secondIndex, 0, identity(target.first));
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  const reopenedOwner = target.group
    ? reopened.resolve(target.group.id)
    : reopened.slides.items[target.slide.index].elements;
  const after = (target.group ? reopenedOwner.children : reopenedOwner.items).map(identity);
  if (JSON.stringify(after) !== JSON.stringify(expected)) {
    throw new Error(`Z-order edit did not survive re-import for ${target.group ? `group ${target.group.id}` : `slide ${target.slide.index + 1}`}.`);
  }
  const changedParts = await changedPackageParts(bytes, output.bytes);
  const expectedPart = `ppt/slides/slide${target.slide.index + 1}.xml`;
  if (changedParts.length !== 1 || changedParts[0] !== expectedPart) {
    throw new Error(`Z-order edit changed unexpected parts for ${target.group ? `group ${target.group.id}` : `slide ${target.slide.index + 1}`}: ${changedParts.join(", ")}`);
  }
  return { status: "passed", scope: target.group ? "group" : "slide", slide: target.slide.index + 1, ...(target.group ? { groupId: target.group.id } : {}), movedId: target.first.id, before, after, changedParts };
}

async function verifyOneSlideReuse(bytes) {
  const presentation = await importPresentation(bytes);
  const sourceSlide = presentation.slides.items[0];
  if (!sourceSlide.cloneCapability.supported) return { status: "blocked", reason: sourceSlide.cloneCapability.blockedReason || "source slide is not reusable" };
  const originalSlideCount = presentation.slides.count;
  const allSlidesCloneCapable = presentation.slides.items.every((slide) => slide.cloneCapability.supported);
  sourceSlide.duplicate();
  const output = await PresentationFile.exportPptx(presentation);
  const reopened = await importPresentation(output.bytes);
  if (reopened.slides.count !== originalSlideCount + 1) throw new Error("Source slide reuse produced an unexpected slide count.");
  const sourceZip = await JSZip.loadAsync(bytes);
  const outputZip = await JSZip.loadAsync(output.bytes);
  const sourceSlideXml = await sourceZip.file("ppt/slides/slide1.xml").async("uint8array");
  const outputSlideXml = await outputZip.file("ppt/slides/slide1.xml").async("uint8array");
  if (!Buffer.from(sourceSlideXml).equals(Buffer.from(outputSlideXml))) throw new Error("Source slide changed during reuse.");
  return {
    status: "passed",
    sourceSlideId: sourceSlide.id,
    sourceSlideUnchanged: true,
    outputSlideCount: reopened.slides.count,
    allSlidesCloneCapable,
  };
}

async function verifySourceComponentReuse(bytes) {
  const presentation = await importPresentation(bytes);
  const candidates = parseNdjson(presentation.inspect({ kind: "componentCandidate", maxChars: Infinity }).ndjson)
    .filter((candidate) => candidate.status === "inspect-only" && candidate.occurrences?.some((occurrence) => occurrence.reuseCapability?.supported === true));
  if (!candidates.length) return { status: "blocked", reason: "no closed source component candidate was discovered" };
  const originalSlideCount = presentation.slides.count;
  let lastError;
  for (const candidate of candidates) {
    const occurrenceIndex = candidate.occurrences.findIndex((occurrence) => occurrence.reuseCapability?.supported === true);
    const occurrence = candidate.occurrences[occurrenceIndex];
    const sourceSlideIndex = presentation.slides.items.findIndex((slide) => slide.id === occurrence.slideId);
    if (sourceSlideIndex < 0) continue;
    try {
      const sourceZip = await JSZip.loadAsync(bytes);
      const sourceSlidePart = `ppt/slides/slide${sourceSlideIndex + 1}.xml`;
      const sourceSlideXml = await sourceZip.file(sourceSlidePart).async("uint8array");
      const clone = presentation.reuseSourceComponent({
        candidateId: candidate.candidateId,
        occurrenceIndex,
        expectedCandidate: candidate,
      });
      const output = await PresentationFile.exportPptx(presentation);
      const reopened = await importPresentation(output.bytes);
      if (reopened.slides.count !== originalSlideCount + 1 || clone.elements.count !== 1) {
        throw new Error(`Source component reuse produced an unexpected clone shape for ${candidate.candidateId}.`);
      }
      const outputZip = await JSZip.loadAsync(output.bytes);
      const outputSlideXml = await outputZip.file(sourceSlidePart).async("uint8array");
      if (!Buffer.from(sourceSlideXml).equals(Buffer.from(outputSlideXml))) {
        throw new Error(`Source component reuse changed its source slide for ${candidate.candidateId}.`);
      }
      return {
        status: "passed",
        candidateId: candidate.candidateId,
        occurrenceIndex,
        sourceSlide: sourceSlideIndex + 1,
        cloneSlide: clone.index + 1,
        cloneElements: clone.elements.count,
        reopenedSlides: reopened.slides.count,
        sourceSlideUnchanged: true,
      };
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError || new Error("Source component reuse candidates could not be executed.");
}

function firstTextRun(presentation) {
  for (const slide of presentation.slides.items) {
    const shapes = [];
    const collect = (group) => {
      shapes.push(...(group.shapes?.items || []));
      for (const child of group.groups?.items || []) collect(child);
    };
    shapes.push(...(slide.shapes?.items || []));
    for (const group of slide.groups?.items || []) collect(group);
    for (const shape of shapes) {
      for (const paragraph of shape.text?.paragraphs || []) {
        const run = paragraph.runs?.find((candidate) => typeof candidate.text === "string" && candidate.text.trim().length >= 4);
        if (run) return { slide: slide.index + 1, shape, run };
      }
    }
  }
  return undefined;
}

function firstPlacementObject(presentation) {
  for (const slide of presentation.slides.items) {
    const object = (slide.nativeObjects?.items || []).find((candidate) => candidate.placementCapability?.supported === true);
    if (object) return { slide: slide.index + 1, object };
  }
  return undefined;
}

function firstReplaceableImagePair(presentation) {
  const byType = new Map();
  for (const slide of presentation.slides.items) {
    for (const image of slide.images?.items || []) {
      const dataUrl = image.dataUrl;
      const contentType = dataUrl?.match(/^data:([^;]+);base64,/u)?.[1]?.toLowerCase();
      if (!contentType) continue;
      const candidates = byType.get(contentType) || [];
      if (!candidates.some((candidate) => candidate.image.dataUrl === dataUrl)) candidates.push({ slide, image });
      byType.set(contentType, candidates);
    }
  }
  for (const candidates of byType.values()) {
    if (candidates.length >= 2) return {
      target: candidates[0].image,
      replacement: candidates[1].image,
    };
  }
  return undefined;
}

function firstTableCell(presentation) {
  for (const slide of presentation.slides.items) {
    for (const table of slide.tables?.items || []) {
      for (let row = 0; row < table.rows; row += 1) {
        for (let column = 0; column < table.columns; column += 1) {
          const value = String(table.values[row]?.[column] ?? "");
          if (value.length > 0) return { slide, table, row, column, value };
        }
      }
    }
  }
  return undefined;
}

async function packageEvidence(bytes) {
  const zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
  const slides = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide[1-9][0-9]*[.]xml$/u.test(name))
    .sort((left, right) => slideOrdinal(left) - slideOrdinal(right));
  let rawObjectCount = 0;
  let nestedObjectCount = 0;
  const rawRootKinds = {};
  for (const slide of slides) {
    const roots = directPresentationChildren(await zip.file(slide).async("text"), "spTree");
    for (const root of roots) {
      if (["nvGrpSpPr", "grpSpPr", "extLst"].includes(root.localName)) continue;
      rawObjectCount += 1;
      if (root.localName === "grpSp") nestedObjectCount += countNestedObjects(root.xml, "grpSp");
      rawRootKinds[root.localName] = (rawRootKinds[root.localName] || 0) + 1;
    }
  }
  return {
    slideCount: slides.length,
    rawObjectCount,
    nestedObjectCount,
    visibleObjectCount: rawObjectCount + nestedObjectCount,
    rawRootKinds: sortObject(rawRootKinds),
  };
}

function countNestedObjects(xml, parentLocalName) {
  return directPresentationChildren(xml, parentLocalName)
    .filter((child) => !["nvGrpSpPr", "grpSpPr", "extLst"].includes(child.localName))
    .reduce((count, child) => count + 1 + (child.localName === "grpSp" ? countNestedObjects(child.xml, "grpSp") : 0), 0);
}

async function changedPackageParts(sourceBytes, outputBytes) {
  const source = await JSZip.loadAsync(sourceBytes, { checkCRC32: true });
  const output = await JSZip.loadAsync(outputBytes, { checkCRC32: true });
  const names = [...new Set([
    ...Object.keys(source.files).filter((name) => !source.files[name].dir),
    ...Object.keys(output.files).filter((name) => !output.files[name].dir),
  ])].sort();
  const changed = [];
  for (const name of names) {
    const beforeFile = source.file(name);
    const afterFile = output.file(name);
    if (!beforeFile || !afterFile) {
      changed.push(name);
      continue;
    }
    const before = await beforeFile.async("uint8array");
    const after = await afterFile.async("uint8array");
    if (!before || !after || !Buffer.from(before).equals(Buffer.from(after))) changed.push(name);
  }
  return changed;
}

function dataUrlBytes(dataUrl) {
  const match = /^data:[^;]+;base64,([A-Za-z0-9+/=\s]+)$/u.exec(String(dataUrl || ""));
  if (!match) throw new Error("Expected a base64 image data URL.");
  return Buffer.from(match[1].replace(/\s/gu, ""), "base64");
}

function profileSummary(profile) {
  return {
    schema: profile.schema,
    sourceBound: profile.source?.sourceBound === true,
    revisionSha256: profile.source?.revisionSha256,
    canvas: profile.canvas,
    layoutFamilies: profile.layoutFamilies?.length || 0,
    slideArchetypes: profile.slideArchetypes?.length || 0,
    reusableComponents: profile.reusableComponents?.length || 0,
    componentCandidates: Number(profile.componentCandidates?.total || 0),
    sourceTheme: profile.designLanguage?.theme?.sourceBound === true
      ? {
        id: profile.designLanguage.theme.id,
        name: profile.designLanguage.theme.name,
        colorSchemeName: profile.designLanguage.theme.colorSchemeName,
        xmlSha256: profile.designLanguage.theme.xmlSha256,
      }
      : null,
    svgAssets: profile.designLanguage?.vectorAssets?.assetCount || 0,
    nativeOpaque: profile.nativeOpaque?.count || profile.nativeOpaque?.length || 0,
  };
}

async function importPresentation(bytes) {
  return PresentationFile.importPptx(new FileBlob(bytes, { type: PPTX_MIME }));
}

async function readPptx(filePath) {
  const info = await stat(filePath);
  if (!info.isFile() || info.size < 1 || info.size > MAX_SOURCE_BYTES) throw new RangeError(`PPTX input is outside 1..${MAX_SOURCE_BYTES}: ${filePath}`);
  const bytes = await readFile(filePath);
  if (bytes.byteLength !== info.size) throw new Error(`PPTX input changed while reading: ${filePath}`);
  return bytes;
}

function parseNdjson(value) {
  return String(value || "").split("\n").filter(Boolean).map((line) => JSON.parse(line));
}

function counts(values) {
  return sortObject(values.reduce((result, value) => {
    result[value] = (result[value] || 0) + 1;
    return result;
  }, {}));
}

function sortObject(value) {
  return Object.fromEntries(Object.entries(value).sort(([left], [right]) => left.localeCompare(right)));
}

function slideOrdinal(name) {
  return Number(/slide([1-9][0-9]*)[.]xml$/u.exec(name)?.[1]);
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function parseArgs(argv) {
  let assetsDir = DEFAULT_ASSETS_DIR;
  let output = DEFAULT_OUTPUT;
  for (let index = 0; index < argv.length; index += 1) {
    if (argv[index] === "--assets-dir") assetsDir = argv[++index];
    else if (argv[index] === "--output") output = argv[++index];
    else throw new Error(`Unknown option ${argv[index]}.`);
  }
  return { assetsDir, output };
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const evidence = await collectSixSampleEvidence(options);
  const output = path.resolve(options.output);
  await mkdir(path.dirname(output), { recursive: true });
  await writeFile(output, `${JSON.stringify(evidence, null, 2)}\n`);
  process.stdout.write(`${JSON.stringify({ ok: true, output, ...evidence.totals })}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
