#!/usr/bin/env node

import { createHash } from "node:crypto";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

import { FileBlob, PresentationFile } from "../src/index.mjs";
import { SOURCES } from "./pptx-source-reuse-benchmark.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const EVIDENCE_SCHEMA = "office-kit/pptx-source-continuation-evidence/v1";
const TOPOLOGY_PARTS = new Set(["[Content_Types].xml", "ppt/_rels/presentation.xml.rels", "ppt/presentation.xml"]);

// The operation is intentionally small: re-open the newly cloned slide, then
// edit one existing typed leaf.  The SVG case changes the image payload while
// retaining the original image element and content type; it exercises a
// different source-derived editing boundary without adding topology.
const CONTINUATION_KIND = Object.freeze({
  "suanzhi-future-2026": "text",
  "blue-gray-acid-template": "text",
  "mckinsey-customer-loyalty": "svg-image",
});

export async function runSourceContinuationBenchmark(assetsDir) {
  const results = [];
  for (const source of SOURCES) {
    const inputPath = path.join(assetsDir, source.fileName);
    const sourceBytes = await readFile(inputPath);
    const sourceZip = await JSZip.loadAsync(sourceBytes);
    const sourceSlidePath = `ppt/slides/slide${source.slide}.xml`;
    const sourceSlideBytes = await sourceZip.file(sourceSlidePath).async("uint8array");
    const presentation = await PresentationFile.importPptx(new FileBlob(sourceBytes, { type: PPTX_MIME }));
    const sourceSlideCount = presentation.slides.count;
    const origin = presentation.slides.items[source.slide - 1];
    if (!origin.cloneCapability.supported) throw new Error(`${source.id} cannot run continuation: ${origin.cloneCapability.blockedReason}`);
    const pendingClone = origin.duplicate();
    // Append the derived page rather than inserting it next to the source. This
    // keeps dynamic slide-number fields on every original page unchanged, so
    // the native render oracle measures actual collateral drift instead of the
    // intentional renumbering caused by an insertion in the middle.
    pendingClone.moveTo(sourceSlideCount);
    const clonedOutput = await PresentationFile.exportPptx(presentation);
    const reopenedClone = await PresentationFile.importPptx(clonedOutput.bytes);
    const clone = reopenedClone.slides.items[sourceSlideCount];
    const kind = CONTINUATION_KIND[source.id];
    const target = applyContinuation(clone, kind);
    const output = await PresentationFile.exportPptx(reopenedClone);
    const verified = await PresentationFile.importPptx(output.bytes);
    const verifiedClone = verified.slides.items[sourceSlideCount];
    const verifiedTarget = verifyContinuation(verifiedClone, target);
    const outputZip = await JSZip.loadAsync(output.bytes);
    const sourcePartNames = Object.keys(sourceZip.files).filter((name) => !sourceZip.files[name].dir).sort();
    const changedExistingParts = [];
    for (const name of sourcePartNames) {
      const before = await sourceZip.file(name).async("uint8array");
      const afterFile = outputZip.file(name);
      if (!afterFile) {
        changedExistingParts.push(name);
        continue;
      }
      const after = await afterFile.async("uint8array");
      if (!Buffer.from(before).equals(Buffer.from(after))) changedExistingParts.push(name);
    }
    const outputPartNames = Object.keys(outputZip.files).filter((name) => !outputZip.files[name].dir).sort();
    const result = {
      id: source.id,
      fileName: source.fileName,
      sourceSha256: sha256(sourceBytes),
      sourceSlideSha256: sha256(sourceSlideBytes),
      sourceSlidePart: sourceSlidePath,
      sourceSlideCount,
      outputSlideCount: verified.slides.count,
      kind,
      cloneOutputSha256: sha256(clonedOutput.bytes),
      outputSha256: sha256(output.bytes),
      sourceSlideUnchanged: Buffer.from(sourceSlideBytes).equals(Buffer.from(await outputZip.file(sourceSlidePath).async("uint8array"))),
      changedExistingParts,
      topologyChangedParts: changedExistingParts.filter((name) => TOPOLOGY_PARTS.has(name)),
      nonTopologyChangedParts: changedExistingParts.filter((name) => !TOPOLOGY_PARTS.has(name)),
      addedParts: outputPartNames.filter((name) => !sourceZip.file(name)),
      target,
      verifiedTarget,
    };
    if (!result.sourceSlideUnchanged || result.nonTopologyChangedParts.length > 0 || result.outputSlideCount !== result.sourceSlideCount + 1) {
      throw new Error(`${source.id} continuation changed a non-target source part or produced the wrong slide count.`);
    }
    results.push(result);
  }
  return { schema: EVIDENCE_SCHEMA, sources: results };
}

export function applyContinuation(slide, kind) {
  if (kind === "text") {
    const shape = slide.shapes.items.find((candidate) => candidate.text?.value);
    if (!shape) throw new Error(`No editable text leaf found on cloned slide ${slide.index + 1}.`);
    const before = shape.text.value;
    const after = `${before} · OfficeKit continuation`;
    shape.text.set(after);
    return { kind, id: shape.id, before, after };
  }
  if (kind === "svg-image") {
    const image = slide.images.items.find((candidate) => candidate.dataUrl?.startsWith("data:image/svg+xml;base64,"));
    if (!image) throw new Error(`No SVG image found on cloned slide ${slide.index + 1}.`);
    const before = Buffer.from(image.dataUrl.split(",", 2)[1], "base64").toString("utf8");
    const after = before.replace(/^<svg\b/iu, '<svg data-officekit="continuation"');
    if (after === before) throw new Error("SVG continuation marker could not be inserted.");
    image.replace({ dataUrl: `data:image/svg+xml;base64,${Buffer.from(after).toString("base64")}` });
    return { kind, id: image.id, marker: "data-officekit=\"continuation\"" };
  }
  throw new Error(`Unknown continuation kind ${kind}.`);
}

export function verifyContinuation(slide, target) {
  if (target.kind === "text") {
    const shape = slide.shapes.getItem(target.id) || slide.shapes.items.find((candidate) => candidate.text?.value === target.after);
    if (!shape || shape.text.value !== target.after) throw new Error(`Continuation text ${target.id} did not survive reimport.`);
    return { id: shape.id, value: shape.text.value };
  }
  const image = slide.images.getItem?.(target.id) || slide.images.items.find((candidate) => candidate.dataUrl?.startsWith("data:image/svg+xml;base64,"));
  if (!image || !Buffer.from(image.dataUrl.split(",", 2)[1], "base64").toString("utf8").includes(target.marker)) {
    throw new Error(`Continuation SVG image ${target.id} did not survive reimport.`);
  }
  return { id: image.id, marker: target.marker };
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
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
  if (!assetsDir || !output) throw new Error("Usage: pptx-source-continuation-benchmark.mjs --assets-dir <dir> --output <evidence.json> [--force]");
  return { assetsDir: path.resolve(assetsDir), output: path.resolve(output), force };
}

async function main() {
  const { assetsDir, output, force } = parseArgs(process.argv.slice(2));
  await mkdtemp(path.join(os.tmpdir(), "officekit-pptx-source-continuation-"));
  const evidence = await runSourceContinuationBenchmark(assetsDir);
  await writeFile(output, `${JSON.stringify(evidence, null, 2)}\n`, { flag: force ? "w" : "wx" });
  process.stdout.write(`${JSON.stringify({ ok: true, output, sources: evidence.sources.length })}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
