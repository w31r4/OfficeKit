#!/usr/bin/env node

import { createHash } from "node:crypto";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

import { FileBlob, PresentationFile } from "../src/index.mjs";
import { SOURCES } from "./pptx-source-reuse-benchmark.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const EVIDENCE_SCHEMA = "office-kit/pptx-source-continuation-evidence/v2";
const TOPOLOGY_PARTS = new Set(["[Content_Types].xml", "ppt/_rels/presentation.xml.rels", "ppt/presentation.xml"]);
const OVERLAY_IMAGE = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";
const OVERLAY_NAMES = Object.freeze({
  text: "officekit-source-derived-text",
  accent: "officekit-source-derived-accent",
  image: "officekit-source-derived-image",
});

const CONTINUATION_KIND = Object.freeze({
  "suanzhi-future-2026": "bounded-overlay",
  "blue-gray-acid-template": "bounded-overlay",
  "mckinsey-customer-loyalty": "bounded-overlay",
});

export async function runSourceContinuationBenchmark(assetsDir, options = {}) {
  const repetitions = options.repetitions ?? 3;
  const selectedSources = options.sourceIds?.length
    ? SOURCES.filter((source) => options.sourceIds.includes(source.id))
    : SOURCES;
  const results = [];
  for (const source of selectedSources) {
    const inputPath = path.join(assetsDir, source.fileName);
    const sourceBytes = await readFile(inputPath);
    const runs = [];
    for (let repetition = 1; repetition <= repetitions; repetition += 1) {
      options.onProgress?.({ sourceId: source.id, repetition, repetitions });
      runs.push(await runSourceContinuationCase(source, sourceBytes, repetition));
    }
    const outputHashes = runs.map(({ outputSha256 }) => outputSha256);
    const canonicalHashes = runs.map(({ canonicalOpcSha256 }) => canonicalOpcSha256);
    const cloneHashes = runs.map(({ cloneOutputSha256 }) => cloneOutputSha256);
    const footprints = runs.map(({ overlayChangedExistingParts, overlayAddedParts, overlayRemovedParts }) =>
      JSON.stringify({ overlayChangedExistingParts, overlayAddedParts, overlayRemovedParts }));
    if (new Set(outputHashes).size !== 1 || new Set(cloneHashes).size !== 1 || new Set(footprints).size !== 1) {
      throw new Error(`${source.id} continuation output or mutation footprint was not deterministic: ${JSON.stringify({ outputHashes, canonicalHashes, cloneHashes, footprints, zipEntryMetadata: runs.map(({ zipEntryMetadata }) => zipEntryMetadata) })}`);
    }
    results.push({
      ...runs[0],
      repetitions: runs.length,
      deterministic: true,
      repeatOutputSha256s: outputHashes,
      repeatCanonicalOpcSha256s: canonicalHashes,
      repeatCloneOutputSha256s: cloneHashes,
      repeatMutationFootprints: footprints,
    });
  }
  return { schema: EVIDENCE_SCHEMA, sources: results };
}

async function runSourceContinuationCase(source, sourceBytes, repetition) {
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
  const cloneZip = await JSZip.loadAsync(clonedOutput.bytes);
  const clonedSlidePath = await presentationSlidePart(cloneZip, sourceSlideCount);
  const clonedSlideXml = await cloneZip.file(clonedSlidePath).async("text");
  const reopenedClone = await PresentationFile.importPptx(clonedOutput.bytes);
  const clone = reopenedClone.slides.items[sourceSlideCount];
  const kind = CONTINUATION_KIND[source.id];
  const target = applyContinuation(clone, kind);
  const output = await PresentationFile.exportPptx(reopenedClone);
  const verified = await PresentationFile.importPptx(output.bytes);
  const verifiedClone = verified.slides.items[sourceSlideCount];
  const verifiedTarget = verifyContinuation(verifiedClone, target);
  const outputZip = await JSZip.loadAsync(output.bytes);
  const outputClonedSlidePath = await presentationSlidePart(outputZip, sourceSlideCount);
  if (outputClonedSlidePath !== clonedSlidePath) throw new Error(`${source.id} cloned SlidePart identity drifted during continuation.`);
  const outputClonedSlideXml = await outputZip.file(clonedSlidePath).async("text");
  const maskedClonedSlideXml = removeOverlayElements(outputClonedSlideXml);
  const overlayPartDiff = await diffPackages(cloneZip, outputZip);
  const clonedSlideRels = relationshipPartPath(clonedSlidePath);
  const allowedOverlayChanges = new Set(["[Content_Types].xml", clonedSlidePath, clonedSlideRels]);
  const unexpectedOverlayChanges = overlayPartDiff.changedExistingParts.filter((name) => !allowedOverlayChanges.has(name));
  const addedMediaParts = overlayPartDiff.addedParts.filter((name) => /^ppt\/media\/[^/]+\.png$/u.test(name));
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
    repetition,
    sourceSha256: sha256(sourceBytes),
    sourceSlideSha256: sha256(sourceSlideBytes),
    sourceSlidePart: sourceSlidePath,
    sourceSlideCount,
    outputSlideCount: verified.slides.count,
    kind,
    cloneOutputSha256: sha256(clonedOutput.bytes),
    outputSha256: sha256(output.bytes),
    canonicalOpcSha256: await canonicalOpcSha256(outputZip),
    sourceSlideUnchanged: Buffer.from(sourceSlideBytes).equals(Buffer.from(await outputZip.file(sourceSlidePath).async("uint8array"))),
    changedExistingParts,
    topologyChangedParts: changedExistingParts.filter((name) => TOPOLOGY_PARTS.has(name)),
    nonTopologyChangedParts: changedExistingParts.filter((name) => !TOPOLOGY_PARTS.has(name)),
    addedParts: outputPartNames.filter((name) => !sourceZip.file(name)),
    clonedSlidePart: clonedSlidePath,
    targetXmlMaskedEqual: maskedClonedSlideXml === clonedSlideXml,
    overlayChangedExistingParts: overlayPartDiff.changedExistingParts,
    overlayAddedParts: overlayPartDiff.addedParts,
    overlayRemovedParts: overlayPartDiff.removedParts,
    unexpectedOverlayChanges,
    addedMediaParts,
    zipEntryMetadata: Object.fromEntries(
      [...overlayPartDiff.changedExistingParts, ...overlayPartDiff.addedParts]
        .sort()
        .map((name) => [name, zipMetadata(outputZip.files[name])]),
    ),
    target,
    verifiedTarget,
  };
  if (!result.sourceSlideUnchanged || result.nonTopologyChangedParts.length > 0 || result.outputSlideCount !== result.sourceSlideCount + 1 ||
      !result.targetXmlMaskedEqual || result.unexpectedOverlayChanges.length > 0 || result.overlayRemovedParts.length > 0 || result.addedMediaParts.length !== 1) {
    throw new Error(`${source.id} continuation changed a non-target source part or produced the wrong slide count.`);
  }
  return result;
}

export function applyContinuation(slide, kind) {
  if (kind !== "bounded-overlay") throw new Error(`Unknown continuation kind ${kind}.`);
  const capability = slide.continuationCapability;
  if (!capability?.ready || capability.profile !== "bounded-overlay" || capability.embeddedImage !== true) {
    throw new Error(`Slide ${slide.index + 1} is not ready for a bounded source-derived overlay.`);
  }
  const text = `OfficeKit continuation · slide ${slide.index + 1}`;
  const textShape = slide.shapes.add({
    name: OVERLAY_NAMES.text,
    geometry: "textbox",
    position: { left: 56, top: 28, width: 360, height: 52 },
    fill: "#0F172A",
    line: { fill: "#0F172A", width: 0 },
    text,
    textStyle: { fontFamily: "Arial", fontSize: 18, bold: true, color: "#FFFFFF" },
    accessibility: { title: "OfficeKit continuation marker" },
  });
  const accent = slide.shapes.add({
    name: OVERLAY_NAMES.accent,
    geometry: "ellipse",
    position: { left: 424, top: 36, width: 36, height: 36 },
    fill: "#F97316",
    line: { fill: "#C2410C", width: 1 },
    accessibility: { decorative: true },
  });
  const image = slide.images.add({
    name: OVERLAY_NAMES.image,
    alt: "Source-derived continuation image",
    dataUrl: OVERLAY_IMAGE,
    fit: "stretch",
    position: { left: 468, top: 36, width: 36, height: 36 },
  });
  return {
    kind,
    capability,
    text: { id: textShape.id, name: OVERLAY_NAMES.text, value: text, position: textShape.position },
    accent: { id: accent.id, name: OVERLAY_NAMES.accent, geometry: accent.geometry, position: accent.position },
    image: { id: image.id, name: OVERLAY_NAMES.image, alt: image.alt, position: image.position, sha256: sha256(Buffer.from(OVERLAY_IMAGE.split(",", 2)[1], "base64")) },
  };
}

export function verifyContinuation(slide, target) {
  if (target.kind !== "bounded-overlay") throw new Error(`Unknown continuation kind ${target.kind}.`);
  const text = slide.shapes.items.find((candidate) => candidate.name === target.text.name);
  const accent = slide.shapes.items.find((candidate) => candidate.name === target.accent.name);
  const image = slide.images.items.find((candidate) => candidate.name === target.image.name);
  const imageSha256 = image?.dataUrl ? sha256(Buffer.from(image.dataUrl.split(",", 2)[1], "base64")) : undefined;
  if (text?.text?.value !== target.text.value || accent?.geometry !== target.accent.geometry || image?.alt !== target.image.alt || imageSha256 !== target.image.sha256) {
    throw new Error(`Bounded continuation overlay did not survive reimport on slide ${slide.index + 1}.`);
  }
  return {
    kind: target.kind,
    text: { id: text.id, name: text.name, value: text.text.value, position: text.position },
    accent: { id: accent.id, name: accent.name, geometry: accent.geometry, position: accent.position },
    image: { id: image.id, name: image.name, alt: image.alt, position: image.position, sha256: imageSha256 },
  };
}

async function presentationSlidePart(zip, slideIndex) {
  const presentationXml = await zip.file("ppt/presentation.xml").async("text");
  const relationshipXml = await zip.file("ppt/_rels/presentation.xml.rels").async("text");
  const slideTags = [...presentationXml.matchAll(/<p:sldId\b[^>]*\/?\s*>/gu)].map((match) => match[0]);
  const slideTag = slideTags[slideIndex];
  const relationshipId = attributeValue(slideTag, "r:id");
  if (!relationshipId) throw new Error(`Cannot resolve slide relationship at index ${slideIndex}.`);
  const relationshipTag = [...relationshipXml.matchAll(/<Relationship\b[^>]*\/?\s*>/gu)]
    .map((match) => match[0])
    .find((tag) => attributeValue(tag, "Id") === relationshipId);
  const target = attributeValue(relationshipTag, "Target");
  if (!target) throw new Error(`Cannot resolve SlidePart for ${relationshipId}.`);
  return target.startsWith("/")
    ? path.posix.normalize(target.replace(/^\/+/, ""))
    : path.posix.normalize(path.posix.join("ppt", target));
}

function attributeValue(tag, name) {
  if (!tag) return undefined;
  const escapedName = name.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
  const match = tag.match(new RegExp(`\\b${escapedName}=(?:\"([^\"]*)\"|'([^']*)')`, "u"));
  return match?.[1] ?? match?.[2];
}

function removeOverlayElements(xml) {
  return xml.replace(/<p:(sp|pic)\b[\s\S]*?<\/p:\1>/gu, (outerXml) =>
    Object.values(OVERLAY_NAMES).some((name) => outerXml.includes(`name="${name}"`)) ? "" : outerXml);
}

async function diffPackages(beforeZip, afterZip) {
  const beforeNames = Object.keys(beforeZip.files).filter((name) => !beforeZip.files[name].dir).sort();
  const afterNames = Object.keys(afterZip.files).filter((name) => !afterZip.files[name].dir).sort();
  const changedExistingParts = [];
  for (const name of beforeNames) {
    const after = afterZip.file(name);
    if (!after) continue;
    const [beforeBytes, afterBytes] = await Promise.all([
      beforeZip.file(name).async("uint8array"),
      after.async("uint8array"),
    ]);
    if (!Buffer.from(beforeBytes).equals(Buffer.from(afterBytes))) changedExistingParts.push(name);
  }
  return {
    changedExistingParts,
    addedParts: afterNames.filter((name) => !beforeZip.file(name)),
    removedParts: beforeNames.filter((name) => !afterZip.file(name)),
  };
}

async function canonicalOpcSha256(zip) {
  const hash = createHash("sha256");
  const names = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  for (const name of names) {
    hash.update(name);
    hash.update("\0");
    hash.update(await zip.file(name).async("uint8array"));
    hash.update("\0");
  }
  return hash.digest("hex");
}

function zipMetadata(entry) {
  return entry ? {
    date: entry.date?.toISOString?.(),
    unixPermissions: entry.unixPermissions ?? null,
    dosPermissions: entry.dosPermissions ?? null,
  } : null;
}

function relationshipPartPath(partPath) {
  const directory = path.posix.dirname(partPath);
  return path.posix.join(directory, "_rels", `${path.posix.basename(partPath)}.rels`);
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function parseArgs(argv) {
  let assetsDir;
  let output;
  let force = false;
  const sourceIds = [];
  for (let index = 0; index < argv.length; index += 1) {
    const flag = argv[index];
    if (flag === "--assets-dir") assetsDir = argv[++index];
    else if (flag === "--output") output = argv[++index];
    else if (flag === "--source") sourceIds.push(argv[++index]);
    else if (flag === "--force") force = true;
    else throw new Error(`Unknown option ${flag}.`);
  }
  if (!assetsDir || !output) throw new Error("Usage: pptx-source-continuation-benchmark.mjs --assets-dir <dir> --output <evidence.json> [--source <id>] [--force]");
  return { assetsDir: path.resolve(assetsDir), output: path.resolve(output), force, sourceIds };
}

async function main() {
  const { assetsDir, output, force, sourceIds } = parseArgs(process.argv.slice(2));
  await mkdtemp(path.join(os.tmpdir(), "officekit-pptx-source-continuation-"));
  const evidence = await runSourceContinuationBenchmark(assetsDir, {
    sourceIds,
    onProgress: (event) => process.stderr.write(`${JSON.stringify({ phase: "continuation", ...event })}\n`),
  });
  await writeFile(output, `${JSON.stringify(evidence, null, 2)}\n`, { flag: force ? "w" : "wx" });
  process.stdout.write(`${JSON.stringify({ ok: true, output, sources: evidence.sources.length })}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
