#!/usr/bin/env node

import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

import { FileBlob, PresentationFile } from "../src/index.mjs";
import { SOURCES } from "./pptx-source-reuse-benchmark.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const EVIDENCE_SCHEMA = "office-kit/pptx-source-component-reuse-evidence/v1";
const TOPOLOGY_PARTS = new Set(["[Content_Types].xml", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"]);

export async function runSourceComponentReuse(assetsDir) {
  const results = [];
  for (const source of SOURCES) {
    const sourceBytes = await readFile(path.join(assetsDir, source.fileName));
    const sourceSha256 = sha256(sourceBytes);
    const sourcePresentation = await PresentationFile.importPptx(new FileBlob(sourceBytes, { type: PPTX_MIME, name: source.fileName }));
    const sourceSlideCount = sourcePresentation.slides.count;
    const records = sourcePresentation.inspect({ includeComponentCandidates: true, maxChars: Infinity }).ndjson
      .split("\n").filter(Boolean).map((line) => JSON.parse(line)).filter((record) => record.kind === "componentCandidate");
    const inspectOnlyCandidates = records.filter((record) => record.status === "inspect-only");
    const candidates = inspectOnlyCandidates.filter((record) =>
      record.occurrences?.some((occurrence) => occurrence.reuseCapability?.supported === true),
    );
    candidates.sort((left, right) => continuationScore(sourcePresentation, right, source.id) - continuationScore(sourcePresentation, left, source.id));
    const preflightBlockedCandidates = inspectOnlyCandidates.filter((record) => !candidates.includes(record));
    const failures = {};
    let passed;
    for (const candidate of candidates) {
      const occurrenceIndex = candidate.occurrences.findIndex((occurrence) =>
        occurrence.reuseCapability?.supported === true,
      );
      if (occurrenceIndex < 0) continue;
      try {
        // Every candidate attempt starts from clean source bytes. A failed
        // continuation must not leave a speculative clone in the next
        // candidate's package.
        const presentation = await PresentationFile.importPptx(new FileBlob(sourceBytes, { type: PPTX_MIME, name: source.fileName }));
        const clone = presentation.reuseSourceComponent({
          candidateId: candidate.candidateId,
          occurrenceIndex,
          expectedCandidate: candidate,
        });
        const output = await PresentationFile.exportPptx(presentation);
        const reopened = await PresentationFile.importPptx(output.bytes);
        const nonTargetPartMismatches = await compareNonTopologyParts(sourceBytes, output.bytes);
        const continuedMutation = await continueClonedComponent({
          sourceBytes,
          componentBytes: output.bytes,
          reopened,
          cloneSlideIndex: clone.index,
          sourceSlidePart: sourceSlidePartForOccurrence(candidate.occurrences[occurrenceIndex]),
        });
        const candidateResult = {
          candidateId: candidate.candidateId,
          occurrenceIndex,
          occurrenceCount: candidate.occurrences.length,
          cloneSlideIndex: clone.index,
          cloneElementCount: directElementCount(clone),
          sourceBytes: sourceBytes.length,
          outputBytes: output.bytes.length,
          reopenedSlideCount: reopened.slides.count,
          nonTargetPartMismatches,
          // ZIP containers may carry producer/runtime timestamps. Keep the
          // raw output hash for diagnostics, but use a canonical hash of
          // sorted OPC file names and bytes for repeatable evidence.
          outputPackageContentSha256: await canonicalPackageContentSha256(output.bytes),
          outputSha256: sha256(output.bytes),
          continuedMutation,
        };
        if (nonTargetPartMismatches.length > 0 || continuedMutation.status !== "passed") {
          const failureCode = nonTargetPartMismatches.length > 0 ? "non_target_part_drift" : "component_continuation_failed";
          failures[failureCode] = (failures[failureCode] || 0) + 1;
          continue;
        }
        passed = { ...candidateResult, status: "passed" };
        break;
      } catch (error) {
        const code = error?.code || error?.name || "unknown";
        failures[code] = (failures[code] || 0) + 1;
      }
    }
    results.push({
      id: source.id,
      sourceSha256,
      sourceSlideCount,
      candidateCount: records.length,
      inspectOnlyCandidateCount: inspectOnlyCandidates.length,
      preflightBlockedCandidateCount: preflightBlockedCandidates.length,
      preflightBlockedReasons: preflightBlockedCandidates.map((candidate) => ({
        candidateId: candidate.candidateId,
        reason: candidate.reuseCapability?.reason || "no occurrence has a supported reuse preflight",
      })),
      status: passed?.status || (passed ? "failed" : "blocked"),
      ...(passed || candidates.length ? {} : { blockedReason: "Every inspect-only candidate failed the occurrence-level reuse preflight." }),
      failures,
      ...(passed || {}),
    });
  }
  return { schema: EVIDENCE_SCHEMA, sources: results };
}

function directElementCount(slide) {
  return [slide.shapes, slide.images, slide.tables, slide.charts, slide.groups, slide.connectors, slide.nativeObjects]
    .reduce((total, collection) => total + collection.items.length, 0);
}

function continuationScore(presentation, candidate, sourceId) {
  const occurrenceIndex = candidate.occurrences.findIndex((occurrence) => occurrence.reuseCapability?.supported === true);
  if (occurrenceIndex < 0) return -1;
  const element = presentation.resolve(candidate.occurrences[occurrenceIndex].targetId);
  if (!element) return -1;
  const kind = continuationKind(element);
  const preferred = {
    "suanzhi-future-2026": "imageFrame",
    "blue-gray-acid-template": "shapeText",
    "mckinsey-customer-loyalty": "svgText",
  }[sourceId];
  if (kind === preferred) return 5;
  if (kind === "svgText" || kind === "shapeText") return 3;
  if (kind === "imageFrame" || kind === "shapeFrame") return 2;
  return 0;
}

function continuationKind(element) {
  const isImage = typeof element?.getSvgTextNodes === "function" && typeof element?.replace === "function";
  const isShape = !isImage && element?.text && typeof element.text.value === "string";
  if (isImage && element.svgTextCapability?.supported === true) return "svgText";
  if (isShape && element.text.value) return "shapeText";
  if (isImage) return "imageFrame";
  if (isShape) return "shapeFrame";
  return undefined;
}

async function compareNonTopologyParts(sourceBytes, outputBytes) {
  const sourceZip = await JSZip.loadAsync(sourceBytes);
  const outputZip = await JSZip.loadAsync(outputBytes);
  const mismatches = [];
  for (const name of Object.keys(sourceZip.files)) {
    if (TOPOLOGY_PARTS.has(name)) continue;
    const sourceFile = sourceZip.file(name);
    // ZIP directory entries are packaging hints, not OPC parts.  JSZip may
    // omit empty directory entries when the cloned package is rebuilt; that
    // must not be reported as content drift.
    if (!sourceFile || sourceFile.dir) continue;
    const outputFile = outputZip.file(name);
    if (!outputFile) {
      mismatches.push(`${name}:missing`);
      continue;
    }
    const sourcePart = await sourceFile.async("uint8array");
    const outputPart = await outputFile.async("uint8array");
    if (sourcePart.length !== outputPart.length || sourcePart.some((value, index) => value !== outputPart[index])) {
      mismatches.push(name);
    }
  }
  return mismatches;
}

async function continueClonedComponent({ sourceBytes, componentBytes, reopened, cloneSlideIndex, sourceSlidePart }) {
  const slide = reopened.slides.items[cloneSlideIndex];
  if (!slide) return { status: "failed", reason: "clone slide was not present after reimport" };
  const elements = directElements(slide);
  if (elements.length !== 1) return { status: "failed", reason: `clone slide exposed ${elements.length} direct elements; bounded continuation requires one` };
  const [element] = elements;
  const isImage = typeof element.getSvgTextNodes === "function" && typeof element.replace === "function";
  const isShape = !isImage && element.text && typeof element.text.value === "string";
  let mode;
  const targetId = element.id;
  let expectedValue;
  let value;
  if (isImage && element.svgTextCapability?.supported === true) {
    const node = element.getSvgTextNodes()[0];
    if (!node) return { status: "failed", reason: "clone SVG image exposed no text leaf" };
    mode = "svgText";
    expectedValue = node.text;
    value = `${node.text} (continued)`;
    element.editSvgText(node.id, { expectedHash: node.expectedHash, value });
  } else if (isShape && element.text.value) {
    mode = "shapeText";
    expectedValue = element.text.value;
    value = `${element.text.value} (continued)`;
    element.text.value = value;
  } else if (isImage || isShape) {
    mode = isImage ? "imageFrame" : "shapeFrame";
    expectedValue = element.position.left;
    value = expectedValue + 1;
    element.position = { ...element.position, left: value };
  } else {
    return { status: "failed", reason: "clone element has no bounded continuation operation" };
  }
  const continued = await PresentationFile.exportPptx(reopened);
  const changedParts = await changedPackageParts(componentBytes, continued.bytes);
  const sourceSlideBefore = await packagePart(sourceBytes, sourceSlidePart);
  const sourceSlideAfter = await packagePart(continued.bytes, sourceSlidePart);
  if (!sourceSlideBefore || !sourceSlideAfter || !sourceSlideBefore.equals(sourceSlideAfter)) {
    return { status: "failed", mode, targetId, reason: "source slide changed during clone continuation", changedParts };
  }
  const expectedChanged = mode === "svgText"
    ? changedParts.length === 3 &&
      changedParts.filter((name) => /^ppt\/slides\/slide\d+\.xml$/u.test(name)).length === 1 &&
      changedParts.filter((name) => /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/u.test(name)).length === 1 &&
      changedParts.filter((name) => /^ppt\/media\/[^/]+$/u.test(name)).length === 1
    : changedParts.length === 1 && /^ppt\/slides\/slide\d+\.xml$/u.test(changedParts[0]);
  if (!expectedChanged) return { status: "failed", mode, targetId, expectedValue, value, changedParts, reason: "continuation changed an unexpected package footprint" };
  const verified = await PresentationFile.importPptx(continued.bytes);
  const verifiedElement = directElements(verified.slides.items[cloneSlideIndex])[0];
  const verifiedValue = mode === "svgText"
    ? verifiedElement.getSvgTextNodes()[0]?.text
    : mode === "shapeText" ? verifiedElement.text.value : verifiedElement.position.left;
  if (verifiedValue !== value) return { status: "failed", mode, targetId, expectedValue, value, verifiedValue, changedParts, reason: "continuation did not survive second import" };
  return { status: "passed", mode, targetId, expectedValue, value, verifiedValue, changedParts, sourceSlidePartUnchanged: true };
}

function directElements(slide) {
  return [
    ...(slide?.shapes?.items || []),
    ...(slide?.tables?.items || []),
    ...(slide?.charts?.items || []),
    ...(slide?.images?.items || []),
    ...(slide?.groups?.items || []),
    ...(slide?.nativeObjects?.items || []),
    ...(slide?.connectors?.items || []),
  ];
}

function sourceSlidePartForOccurrence(occurrence) {
  const slideNumber = String(occurrence?.slideId || "").split("/")[2];
  if (!/^[1-9]\d*$/u.test(slideNumber)) throw new Error("Component occurrence did not expose a source slide ID.");
  return `ppt/slides/slide${slideNumber}.xml`;
}

async function packagePart(bytes, name) {
  const zip = await JSZip.loadAsync(bytes);
  const file = zip.file(name);
  return file ? Buffer.from(await file.async("uint8array")) : undefined;
}

async function changedPackageParts(beforeBytes, afterBytes) {
  const before = await JSZip.loadAsync(beforeBytes);
  const after = await JSZip.loadAsync(afterBytes);
  const names = new Set([...Object.keys(before.files), ...Object.keys(after.files)]);
  const changed = [];
  for (const name of names) {
    if (before.files[name]?.dir || after.files[name]?.dir) continue;
    const left = before.file(name);
    const right = after.file(name);
    if (!left || !right) {
      changed.push(name);
      continue;
    }
    const leftBytes = await left.async("uint8array");
    const rightBytes = await right.async("uint8array");
    if (leftBytes.length !== rightBytes.length || leftBytes.some((value, index) => value !== rightBytes[index])) changed.push(name);
  }
  return changed.sort();
}

async function canonicalPackageContentSha256(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const entries = Object.values(zip.files)
    .filter((file) => !file.dir)
    .sort((left, right) => left.name.localeCompare(right.name));
  const hash = createHash("sha256");
  for (const entry of entries) {
    const content = await entry.async("uint8array");
    hash.update(entry.name, "utf8");
    hash.update("\0", "utf8");
    hash.update(String(content.length), "ascii");
    hash.update("\0", "utf8");
    hash.update(content);
    hash.update("\0", "utf8");
  }
  return hash.digest("hex");
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
  if (!assetsDir || !output) throw new Error("Usage: pptx-source-component-reuse.mjs --assets-dir <dir> --output <evidence.json> [--force]");
  return { assetsDir: path.resolve(assetsDir), output: path.resolve(output), force };
}

async function main() {
  const { assetsDir, output, force } = parseArgs(process.argv.slice(2));
  const evidence = await runSourceComponentReuse(assetsDir);
  await writeFile(output, `${JSON.stringify(evidence, null, 2)}\n`, { flag: force ? "w" : "wx" });
  process.stdout.write(`${JSON.stringify({ ok: true, output, sources: evidence.sources.length })}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
