#!/usr/bin/env node

import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

import { PresentationFile } from "../src/presentation/index.mjs";
import { FileBlob } from "../src/shared/file-blob.mjs";
import { SOURCES } from "./pptx-source-reuse-benchmark.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const EVIDENCE_SCHEMA = "office-kit/pptx-svg-leaf-evidence/v1";
const EDIT_SOURCE_ID = "mckinsey-customer-loyalty";
const EDIT_SLIDE = 3;
const REPEATS = 3;

function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

function collectImages(container, output = []) {
  for (const image of container?.images?.items || []) output.push(image);
  for (const group of container?.groups?.items || []) collectImages(group, output);
  return output;
}

function imageBytes(image) {
  const base64 = String(image.dataUrl || "").split(",", 2)[1];
  return Buffer.from(base64 || "", "base64");
}

function leafKindCounts(leaves) {
  const counts = {};
  for (const leaf of leaves) counts[leaf.leafKind] = (counts[leaf.leafKind] || 0) + 1;
  return Object.fromEntries(Object.entries(counts).sort(([left], [right]) => left.localeCompare(right)));
}

async function packageDiff(sourceZip, outputZip) {
  const sourceParts = Object.keys(sourceZip.files).filter((name) => !sourceZip.files[name].dir).sort();
  const outputParts = Object.keys(outputZip.files).filter((name) => !outputZip.files[name].dir).sort();
  const changedExistingParts = [];
  const deletedParts = [];
  for (const name of sourceParts) {
    const output = outputZip.file(name);
    if (!output) {
      deletedParts.push(name);
      continue;
    }
    const [before, after] = await Promise.all([
      sourceZip.file(name).async("uint8array"),
      output.async("uint8array"),
    ]);
    if (!Buffer.from(before).equals(Buffer.from(after))) changedExistingParts.push(name);
  }
  return {
    changedExistingParts,
    deletedParts,
    addedParts: outputParts.filter((name) => !sourceZip.file(name)),
  };
}

async function canonicalOpcSha256(zip) {
  const hash = createHash("sha256");
  const names = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  for (const name of names) {
    const bytes = Buffer.from(await zip.file(name).async("uint8array"));
    hash.update(`${Buffer.byteLength(name, "utf8")}:`, "utf8");
    hash.update(name, "utf8");
    hash.update(`${bytes.length}:`, "utf8");
    hash.update(bytes);
  }
  return hash.digest("hex");
}

function replacementFor(leaf) {
  if (leaf.leafKind === "svgFillRgb") return leaf.value === "#0F766E" ? "#115E59" : "#0F766E";
  if (leaf.leafKind === "svgStrokeRgb") return leaf.value === "#DC2626" ? "#B91C1C" : "#DC2626";
  if (leaf.leafKind === "svgOpacity") return leaf.value === 0.65 ? 0.35 : 0.65;
  if (leaf.leafKind === "svgTransformScalar" && leaf.component === "angle") return leaf.value === -85 ? -80 : -85;
  throw new Error(`No benchmark replacement for ${leaf.leafKind}:${leaf.component || ""}.`);
}

function issuedLeaf(image, leafKind) {
  const leaves = image.getSvgEditLeaves();
  const leaf = leaves.find((candidate) => candidate.leafKind === leafKind &&
    (leafKind !== "svgTransformScalar" || candidate.component === "angle"));
  if (!leaf) throw new Error(`Benchmark image has no issued ${leafKind} leaf.`);
  return leaf;
}

function scalarMutation(before, after, oldValue, newValue) {
  const oldBytes = Buffer.from(String(oldValue), "utf8");
  const newBytes = Buffer.from(String(newValue), "utf8");
  assert.equal(oldBytes.length, newBytes.length, "benchmark replacements must preserve token length");
  assert.equal(before.length, after.length, "one scalar edit must preserve SVG byte length");
  const candidates = [];
  let offset = before.indexOf(oldBytes);
  while (offset >= 0) {
    if (after.subarray(offset, offset + newBytes.length).equals(newBytes)) {
      const masked = Buffer.from(after);
      oldBytes.copy(masked, offset);
      if (masked.equals(before)) candidates.push(offset);
    }
    offset = before.indexOf(oldBytes, offset + 1);
  }
  assert.equal(candidates.length, 1, "one issued scalar must explain the complete SVG byte diff");
  return {
    offset: candidates[0],
    length: oldBytes.length,
    beforeSha256: sha256(before),
    afterSha256: sha256(after),
  };
}

function editsForImage(image) {
  const edits = [];
  for (const leafKind of ["svgFillRgb", "svgStrokeRgb", "svgOpacity", "svgTransformScalar"]) {
    const leaf = issuedLeaf(image, leafKind);
    const value = replacementFor(leaf);
    const before = imageBytes(image);
    const receipt = image.editSvgLeaf(leaf.id, { expectedHash: leaf.expectedHash, value });
    const after = imageBytes(image);
    edits.push({
      leafKind,
      ...(leaf.component ? { component: leaf.component } : {}),
      leafId: leaf.id,
      expectedHash: leaf.expectedHash,
      oldValue: leaf.value,
      value: receipt.value,
      sourceRevisionSha256: leaf.sourceRevisionSha256,
      mutation: scalarMutation(before, after, leaf.value, receipt.value),
    });
  }
  return edits;
}

async function inspectSource(assetsDir, source) {
  const sourceBytes = await readFile(path.join(assetsDir, source.fileName));
  const presentation = await PresentationFile.importPptx(new FileBlob(sourceBytes, { type: PPTX_MIME, name: source.fileName }));
  const images = presentation.slides.items.flatMap((slide) => collectImages(slide));
  const svgImages = images.filter((image) => image.svgEditCapability.sourceSha256);
  const leaves = svgImages.flatMap((image) => image.getSvgEditLeaves());
  return {
    id: source.id,
    fileName: source.fileName,
    sourceSha256: sha256(sourceBytes),
    slideCount: presentation.slides.count,
    imageCount: images.length,
    svgImageCount: svgImages.length,
    supportedSvgImageCount: svgImages.filter((image) => image.svgEditCapability.supported === true).length,
    svgLeafCount: leaves.length,
    leafKinds: leafKindCounts(leaves),
    revisionBoundLeafCount: leaves.filter((leaf) => leaf.sourceRevisionSha256 === sha256(sourceBytes)).length,
  };
}

async function runOneEdit(assetsDir, source, repeat) {
  const sourceBytes = await readFile(path.join(assetsDir, source.fileName));
  const sourceZip = await JSZip.loadAsync(sourceBytes);
  const presentation = await PresentationFile.importPptx(new FileBlob(sourceBytes, { type: PPTX_MIME, name: source.fileName }));
  const image = collectImages(presentation.slides.items[EDIT_SLIDE - 1])[0];
  if (!image || image.svgEditCapability.supported !== true) throw new Error("McKinsey slide 3 SVG image is not editable.");
  const sourceSvgSha256 = image.svgEditCapability.sourceSha256;
  const edits = editsForImage(image);
  const editedSvgBytes = imageBytes(image);
  const editedSvgSha256 = sha256(editedSvgBytes);
  const output = await PresentationFile.exportPptx(presentation);
  const outputZip = await JSZip.loadAsync(output.bytes);
  const diff = await packageDiff(sourceZip, outputZip);
  const addedSvgParts = diff.addedParts.filter((name) => name.endsWith(".svg"));
  assert.equal(addedSvgParts.length, 1, "the bounded SVG edit must add exactly one replacement SVG part");
  const addedSvgBytes = Buffer.from(await outputZip.file(addedSvgParts[0]).async("uint8array"));
  assert.equal(sha256(addedSvgBytes), editedSvgSha256, "the package replacement part must equal the edited SVG bytes");

  const reopened = await PresentationFile.importPptx(output);
  const reopenedImage = reopened.slides.items
    .flatMap((slide) => collectImages(slide))
    .find((candidate) => candidate.svgEditCapability.sourceSha256 === editedSvgSha256);
  assert.ok(reopenedImage, "edited SVG image must survive package re-import");
  const importObjectNdjson = reopened.inspect({ kind: "importObject", includeImportObjects: true, maxChars: Infinity }).ndjson;
  const importObjects = String(importObjectNdjson || "").trim().split("\n").filter(Boolean).map((line) => JSON.parse(line));
  const importedRecord = importObjects.find((record) => record.targetId === reopenedImage.id);
  assert.ok(importedRecord?.typedOperations?.includes("svg-style"));
  assert.ok(importedRecord?.typedOperations?.includes("svg-transform"));

  return {
    repeat,
    sourceSha256: sha256(sourceBytes),
    sourceSvgSha256,
    editedSvgSha256,
    outputArchiveSha256: sha256(output.bytes),
    canonicalOutputSha256: await canonicalOpcSha256(outputZip),
    editedSlide: EDIT_SLIDE,
    editedImageId: image.id,
    edits,
    changedExistingParts: diff.changedExistingParts,
    deletedParts: diff.deletedParts,
    addedParts: diff.addedParts,
    replacementSvgPart: addedSvgParts[0],
    replacementSvgBytesMatch: addedSvgBytes.equals(editedSvgBytes),
    reimported: Boolean(reopenedImage),
    classifiedOperations: importedRecord.typedOperations,
  };
}

async function runBenchmark(assetsDir) {
  const sources = [];
  for (const source of SOURCES) sources.push(await inspectSource(assetsDir, source));
  const editSource = SOURCES.find((source) => source.id === EDIT_SOURCE_ID);
  if (!editSource) throw new Error(`Missing benchmark source ${EDIT_SOURCE_ID}.`);
  const runs = [];
  for (let repeat = 1; repeat <= REPEATS; repeat += 1) runs.push(await runOneEdit(assetsDir, editSource, repeat));
  assert.equal(new Set(runs.map((run) => run.outputArchiveSha256)).size, 1, "three clean-source edits must have one byte-identical PPTX archive hash");
  assert.equal(new Set(runs.map((run) => run.canonicalOutputSha256)).size, 1, "three clean-source edits must have one canonical OPC hash");
  assert.equal(new Set(runs.map((run) => run.editedSvgSha256)).size, 1, "three clean-source edits must have one edited SVG hash");
  assert.equal(new Set(runs.map((run) => JSON.stringify({
    edits: run.edits,
    changedExistingParts: run.changedExistingParts,
    deletedParts: run.deletedParts,
    addedParts: run.addedParts,
  }))).size, 1, "three clean-source edits must have one mutation footprint");
  return { schema: EVIDENCE_SCHEMA, editSourceId: EDIT_SOURCE_ID, repeats: REPEATS, sources, runs };
}

function parseArgs(argv) {
  let assetsDir;
  let output;
  let force = false;
  for (let index = 0; index < argv.length; index += 1) {
    const flag = argv[index];
    if (flag === "--assets-dir") assetsDir = argv[++index];
    else if (flag === "--output") output = argv[++index];
    else if (flag === "--force") force = true;
    else throw new Error(`Unknown option ${flag}.`);
  }
  if (!assetsDir || !output) throw new Error("Usage: pptx-svg-leaf-benchmark.mjs --assets-dir <dir> --output <evidence.json> [--force]");
  return { assetsDir: path.resolve(assetsDir), output: path.resolve(output), force };
}

async function main() {
  const { assetsDir, output, force } = parseArgs(process.argv.slice(2));
  const evidence = await runBenchmark(assetsDir);
  await writeFile(output, `${JSON.stringify(evidence, null, 2)}\n`, { flag: force ? "w" : "wx" });
  process.stdout.write(`${JSON.stringify({ ok: true, output, sources: evidence.sources.length, runs: evidence.runs.length })}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}

export { runBenchmark };
