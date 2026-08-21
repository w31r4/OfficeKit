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
    const presentation = await PresentationFile.importPptx(new FileBlob(sourceBytes, { type: PPTX_MIME, name: source.fileName }));
    const sourceSlideCount = presentation.slides.count;
    const records = presentation.inspect({ includeComponentCandidates: true, maxChars: Infinity }).ndjson
      .split("\n").filter(Boolean).map((line) => JSON.parse(line)).filter((record) => record.kind === "componentCandidate");
    const inspectOnlyCandidates = records.filter((record) => record.status === "inspect-only");
    const candidates = inspectOnlyCandidates.filter((record) =>
      record.occurrences?.some((occurrence) => occurrence.reuseCapability?.supported === true),
    );
    const preflightBlockedCandidates = inspectOnlyCandidates.filter((record) => !candidates.includes(record));
    const failures = {};
    let passed;
    for (const candidate of candidates) {
      const occurrenceIndex = candidate.occurrences.findIndex((occurrence) =>
        occurrence.reuseCapability?.supported === true,
      );
      if (occurrenceIndex < 0) continue;
      try {
        const clone = presentation.reuseSourceComponent({
          candidateId: candidate.candidateId,
          occurrenceIndex,
          expectedCandidate: candidate,
        });
        const output = await PresentationFile.exportPptx(presentation);
        const reopened = await PresentationFile.importPptx(output.bytes);
        const nonTargetPartMismatches = await compareNonTopologyParts(sourceBytes, output.bytes);
        passed = {
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
        };
        if (nonTargetPartMismatches.length > 0) {
          passed.status = "failed";
          passed.failureCode = "non_target_part_drift";
        } else {
          passed.status = "passed";
        }
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
