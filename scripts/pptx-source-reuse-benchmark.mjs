#!/usr/bin/env node

import { createHash } from "node:crypto";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

import { FileBlob, PresentationFile } from "../src/index.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const EVIDENCE_SCHEMA = "office-kit/pptx-source-reuse-evidence/v1";
const TOPOLOGY_PARTS = new Set(["[Content_Types].xml", "ppt/_rels/presentation.xml.rels", "ppt/presentation.xml"]);
export const SOURCES = Object.freeze([
  {
    id: "suanzhi-future-2026",
    fileName: "b34ddad8cf8b_012_算秩未来2026_0127_极致技术&长期主义.pptx",
    slide: 1,
    expected: "passed",
  },
  {
    id: "blue-gray-acid-template",
    fileName: "template.pptx",
    slide: 1,
    expected: "passed",
  },
  {
    id: "mckinsey-customer-loyalty",
    fileName: "ppt169_麦肯锡风_kimsoong_customer_loyalty.pptx",
    slide: 1,
    expected: "passed",
  },
]);

export async function runSourceReuseBenchmark(assetsDir, outputDir) {
  const results = [];
  await mkdir(outputDir, { recursive: true });
  for (const source of SOURCES) {
    const inputPath = path.join(assetsDir, source.fileName);
    const sourceBytes = await readFile(inputPath);
    const sourceZip = await JSZip.loadAsync(sourceBytes);
    const sourceSlidePath = `ppt/slides/slide${source.slide}.xml`;
    const sourceSlideBytes = await sourceZip.file(sourceSlidePath).async("uint8array");
    const presentation = await PresentationFile.importPptx(new FileBlob(sourceBytes, { type: PPTX_MIME }));
    const slide = presentation.slides.items[source.slide - 1];
    const sourceSlideCount = presentation.slides.count;
    const cloneCapability = slide.cloneCapability;
    const cloneCoverage = {
      totalSlides: sourceSlideCount,
      supportedSlides: presentation.slides.items.filter((candidate) => candidate.cloneCapability.supported).length,
      blockedSlides: presentation.slides.items
        .map((candidate, index) => ({ slide: index + 1, reason: candidate.cloneCapability.blockedReason }))
        .filter(({ reason }) => reason),
    };
    const result = {
      id: source.id,
      fileName: source.fileName,
      slide: source.slide,
      sourceSha256: sha256(sourceBytes),
      sourceSlideSha256: sha256(sourceSlideBytes),
      sourceSlidePart: sourceSlidePath,
      cloneCapability,
      expected: source.expected,
      sourceSlideCount,
      cloneCoverage,
    };
    if (!cloneCapability.supported) {
      result.status = "blocked";
      result.blockedReason = cloneCapability.blockedReason || "unsupported source graph";
      results.push(result);
      continue;
    }
    slide.duplicate();
    const output = await PresentationFile.exportPptx(presentation);
    const outputZip = await JSZip.loadAsync(output.bytes);
    const outputPartNames = Object.keys(outputZip.files).filter((name) => !outputZip.files[name].dir).sort();
    const sourcePartNames = Object.keys(sourceZip.files).filter((name) => !sourceZip.files[name].dir).sort();
    const changedExistingParts = [];
    for (const name of sourcePartNames) {
      if (!outputZip.file(name)) {
        changedExistingParts.push(name);
        continue;
      }
      const before = await sourceZip.file(name).async("uint8array");
      const after = await outputZip.file(name).async("uint8array");
      if (!Buffer.from(before).equals(Buffer.from(after))) changedExistingParts.push(name);
    }
    const addedParts = outputPartNames.filter((name) => !sourceZip.file(name));
    const outputSourceSlideBytes = await outputZip.file(sourceSlidePath).async("uint8array");
    const reopened = await PresentationFile.importPptx(output.bytes);
    result.status = "passed";
    result.outputSha256 = sha256(output.bytes);
    result.outputSlideCount = reopened.slides.count;
    result.sourceSlideUnchanged = Buffer.from(sourceSlideBytes).equals(Buffer.from(outputSourceSlideBytes));
    result.changedExistingParts = changedExistingParts;
    result.topologyChangedParts = changedExistingParts.filter((name) => TOPOLOGY_PARTS.has(name));
    result.nonTopologyChangedParts = changedExistingParts.filter((name) => !TOPOLOGY_PARTS.has(name));
    result.addedParts = addedParts;
    result.allSlidesClone = await cloneAllSlidesTwice(sourceBytes, sourceZip, sourceSlideCount);
    results.push(result);
    await writeFile(path.join(outputDir, `${source.id}-clone.pptx`), output.bytes, { flag: "wx" });
  }
  const evidence = { schema: EVIDENCE_SCHEMA, sources: results };
  for (const result of results) {
    if (result.status !== result.expected) throw new Error(`${result.id} expected ${result.expected} but got ${result.status}`);
    if (result.status === "passed" && (!result.sourceSlideUnchanged || result.outputSlideCount !== result.sourceSlideCount + 1 || result.nonTopologyChangedParts.length > 0)) {
      throw new Error(`${result.id} source-derived clone changed an existing source part or has the wrong slide count.`);
    }
    if (result.cloneCoverage.supportedSlides !== result.cloneCoverage.totalSlides ||
        result.cloneCoverage.blockedSlides.length > 0 ||
        result.allSlidesClone?.status !== "passed" ||
        result.allSlidesClone?.deterministic !== true ||
        result.allSlidesClone?.outputSlideCount !== result.sourceSlideCount * 2 ||
        result.allSlidesClone?.nonTopologyChangedParts.length > 0) {
      throw new Error(`${result.id} did not prove deterministic reuse of every source slide.`);
    }
  }
  return evidence;
}

async function cloneAllSlidesTwice(sourceBytes, sourceZip, sourceSlideCount) {
  const cloneOnce = async () => {
    const presentation = await PresentationFile.importPptx(new FileBlob(sourceBytes, { type: PPTX_MIME }));
    const sourceSlides = [...presentation.slides.items];
    if (sourceSlides.some((slide) => !slide.cloneCapability.supported)) {
      return {
        status: "blocked",
        blockedSlides: sourceSlides
          .map((slide, index) => ({ slide: index + 1, reason: slide.cloneCapability.blockedReason }))
          .filter(({ reason }) => reason),
      };
    }
    for (const sourceSlide of sourceSlides) sourceSlide.duplicate();
    const output = await PresentationFile.exportPptx(presentation);
    const reopened = await PresentationFile.importPptx(output.bytes);
    const outputZip = await JSZip.loadAsync(output.bytes);
    const sourcePartNames = Object.keys(sourceZip.files).filter((name) => !sourceZip.files[name].dir).sort();
    const changedExistingParts = [];
    for (const name of sourcePartNames) {
      const outputPart = outputZip.file(name);
      if (!outputPart) {
        changedExistingParts.push(name);
        continue;
      }
      const before = await sourceZip.file(name).async("uint8array");
      const after = await outputPart.async("uint8array");
      if (!Buffer.from(before).equals(Buffer.from(after))) changedExistingParts.push(name);
    }
    return {
      status: "passed",
      outputBytes: output.bytes,
      outputSha256: sha256(output.bytes),
      outputSlideCount: reopened.slides.count,
      changedExistingParts,
      topologyChangedParts: changedExistingParts.filter((name) => TOPOLOGY_PARTS.has(name)),
      nonTopologyChangedParts: changedExistingParts.filter((name) => !TOPOLOGY_PARTS.has(name)),
    };
  };
  const first = await cloneOnce();
  if (first.status !== "passed") return first;
  const second = await cloneOnce();
  if (second.status !== "passed") return second;
  return {
    status: "passed",
    outputSha256: first.outputSha256,
    deterministic: Buffer.from(first.outputBytes).equals(Buffer.from(second.outputBytes)),
    outputSlideCount: first.outputSlideCount,
    changedExistingParts: first.changedExistingParts,
    topologyChangedParts: first.topologyChangedParts,
    nonTopologyChangedParts: first.nonTopologyChangedParts,
    sourceSlideCount,
  };
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
  if (!assetsDir || !output) throw new Error("Usage: pptx-source-reuse-benchmark.mjs --assets-dir <dir> --output <evidence.json> [--force]");
  return { assetsDir: path.resolve(assetsDir), output: path.resolve(output), force };
}

async function main() {
  const { assetsDir, output, force } = parseArgs(process.argv.slice(2));
  const outputDir = await mkdtemp(path.join(os.tmpdir(), "officekit-pptx-source-reuse-"));
  const evidence = await runSourceReuseBenchmark(assetsDir, outputDir);
  await writeFile(output, `${JSON.stringify(evidence, null, 2)}\n`, { flag: force ? "w" : "wx" });
  process.stdout.write(`${JSON.stringify({ ok: true, output, sources: evidence.sources.length })}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
