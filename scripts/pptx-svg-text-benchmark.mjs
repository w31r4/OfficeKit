#!/usr/bin/env node

import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

import { FileBlob, PresentationFile } from "../src/index.mjs";
import { SOURCES } from "./pptx-source-reuse-benchmark.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const EVIDENCE_SCHEMA = "office-kit/pptx-svg-text-evidence/v1";

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function collectImages(container, output = []) {
  for (const image of container?.images?.items || []) output.push(image);
  for (const group of container?.groups?.items || []) collectImages(group, output);
  return output;
}

function changedExistingParts(sourceZip, outputZip) {
  return Promise.all(Object.keys(sourceZip.files)
    .filter((name) => !sourceZip.files[name].dir)
    .map(async (name) => {
      const sourceFile = sourceZip.file(name);
      const outputFile = outputZip.file(name);
      if (!outputFile) return name;
      const [before, after] = await Promise.all([
        sourceFile.async("uint8array"),
        outputFile.async("uint8array"),
      ]);
      return Buffer.from(before).equals(Buffer.from(after)) ? undefined : name;
    }))
    .then((names) => names.filter(Boolean).sort());
}

async function inspectAndEditSource(assetsDir, source) {
  const inputPath = path.join(assetsDir, source.fileName);
  const sourceBytes = await readFile(inputPath);
  const sourceHash = sha256(sourceBytes);
  const sourceZip = await JSZip.loadAsync(sourceBytes);
  const presentation = await PresentationFile.importPptx(new FileBlob(sourceBytes, { type: PPTX_MIME, name: source.fileName }));
  const images = presentation.slides.items.flatMap((slide) => collectImages(slide).map((image) => ({ image, slide: slide.index + 1 })));
  const svgImages = images.map(({ image, slide }) => ({ image, slide, capability: image.svgTextCapability }))
    .filter(({ capability }) => capability.sourceSha256);
  const svgTextNodeCount = svgImages.reduce((total, { capability }) => total + (capability.nodes?.length || 0), 0);
  const result = {
    id: source.id,
    fileName: source.fileName,
    sourceSha256: sourceHash,
    sourceSlideCount: presentation.slides.count,
    imageCount: images.length,
    svgImageCount: svgImages.length,
    svgTextNodeCount,
    status: "passed",
  };
  const editable = svgImages.find(({ capability }) => capability.supported === true && capability.nodes?.length);
  if (!editable) {
    result.status = "not-applicable";
    result.reason = "No bounded SVG image with directly editable text nodes was found.";
    return result;
  }

  const leaf = editable.capability.nodes[0];
  const replacement = `${leaf.text} (OfficeKit)`;
  const edit = editable.image.editSvgText(leaf.id, { expectedHash: leaf.expectedHash, value: replacement });
  const output = await PresentationFile.exportPptx(presentation);
  const outputZip = await JSZip.loadAsync(output.bytes);
  const changedParts = await changedExistingParts(sourceZip, outputZip);
  const reopened = await PresentationFile.importPptx(output);
  const reopenedImages = reopened.slides.items.flatMap((slide) => collectImages(slide));
  const reopenedNode = reopenedImages.flatMap((image) => image.getSvgTextNodes()).find((node) => node.text === replacement);
  result.editedSlide = editable.slide;
  result.editedImageId = editable.image.id;
  result.editedNodeId = leaf.id;
  result.edit = edit;
  result.reimportedValue = reopenedNode?.text;
  result.reimported = reopenedNode?.text === replacement;
  result.changedExistingParts = changedParts;
  result.addedParts = Object.keys(outputZip.files)
    .filter((name) => !outputZip.files[name].dir && !sourceZip.file(name))
    .sort();
  if (!result.reimported) throw new Error(`${source.id}: SVG text edit did not survive re-import.`);
  return result;
}

async function runBenchmark(assetsDir) {
  const sources = [];
  for (const source of SOURCES) sources.push(await inspectAndEditSource(assetsDir, source));
  return { schema: EVIDENCE_SCHEMA, sources };
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
  if (!assetsDir || !output) throw new Error("Usage: pptx-svg-text-benchmark.mjs --assets-dir <dir> --output <evidence.json> [--force]");
  return { assetsDir: path.resolve(assetsDir), output: path.resolve(output), force };
}

async function main() {
  const { assetsDir, output, force } = parseArgs(process.argv.slice(2));
  const evidence = await runBenchmark(assetsDir);
  await writeFile(output, `${JSON.stringify(evidence, null, 2)}\n`, { flag: force ? "w" : "wx" });
  process.stdout.write(`${JSON.stringify({ ok: true, output, sources: evidence.sources.length })}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}

export { runBenchmark };
