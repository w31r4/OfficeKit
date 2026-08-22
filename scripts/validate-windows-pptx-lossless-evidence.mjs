#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const REPOSITORY_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DEFAULT_MANIFEST = path.join(REPOSITORY_ROOT, "evals/pptx-lossless/manifest.v1.json");
const REQUIRED_SOURCES = ["suanzhi-future-2026", "blue-gray-acid-template", "mckinsey-customer-loyalty"];
const EXPECTED_TOTAL_PAGES = 48;
const REQUIRED_CHECKS = [
  "opened",
  "noRepairPrompt",
  "browsedAllSlides",
  "targetEditVisible",
  "nonTargetPagesPixelIdentical",
  "advancedObjectsPreserved",
  "savedCopy",
  "reopenedCopy",
  "sourceProtected",
  "unsupportedCapabilityFailClosed",
];

export function validateWindowsPptxLosslessEvidence(value, {
  expectedCommit = undefined,
  manifest = undefined,
} = {}) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("evidence must be a JSON object");
  if (value.schema !== "office-kit.windows-pptx-lossless-evidence.v1") throw new Error("unsupported Windows PPTX lossless evidence schema");
  if (value.method !== "human-observed-windows-powerpoint") throw new Error("evidence must come from a human-observed Windows PowerPoint host");
  if (!/^20\d\d-\d\d-\d\dT/.test(String(value.checkedAt || ""))) throw new Error("checkedAt must be an ISO timestamp");
  if (value.host?.platform !== "win32-x64") throw new Error("host.platform must be win32-x64");
  if (value.host?.powerpoint?.installed !== true || !String(value.host?.powerpoint?.version || "")) {
    throw new Error("Microsoft PowerPoint installation and version are required");
  }
  if (!/^20\d\d-\d\d-\d\dT/.test(String(value.host?.observedAt || ""))) throw new Error("host.observedAt must be an ISO timestamp");
  if (!value.visualReview || typeof value.visualReview !== "object") throw new Error("visualReview is required");
  if (!/^20\d\d-\d\d-\d\dT/.test(String(value.visualReview.observedAt || ""))) throw new Error("visualReview.observedAt must be an ISO timestamp");
  if (value.host.observedAt.slice(0, 10) !== value.checkedAt.slice(0, 10)) throw new Error("host.observedAt must be on the evidence date");
  if (value.visualReview.observedAt.slice(0, 10) !== value.checkedAt.slice(0, 10)) throw new Error("visualReview must be on the evidence date");
  if (value.visualReview.renderer !== "Microsoft PowerPoint") throw new Error("visualReview.renderer must be Microsoft PowerPoint");
  if (value.visualReview.pagesCompared !== EXPECTED_TOTAL_PAGES) throw new Error(`visualReview.pagesCompared must cover all ${EXPECTED_TOTAL_PAGES} frozen sample pages`);
  if (!String(value.visualReview.evidencePath || "")) throw new Error("visualReview.evidencePath is required");
  if (!Array.isArray(value.visualReview.pageComparisons) || value.visualReview.pageComparisons.length !== EXPECTED_TOTAL_PAGES) {
    throw new Error(`visualReview.pageComparisons must contain exactly ${EXPECTED_TOTAL_PAGES} page records`);
  }
  if (!/^(?:[0-9a-f]{40}|[0-9a-f]{64})$/u.test(String(value.commit || ""))) throw new Error("commit must be a 40-character SHA-1 or 64-character SHA-256");
  if (expectedCommit && value.commit !== expectedCommit) throw new Error("evidence commit does not match the checked-out commit");
  if (!Array.isArray(value.sources) || value.sources.length !== REQUIRED_SOURCES.length) throw new Error("evidence must contain exactly the three lossless PPTX sources");

  const expectedSources = new Map((manifest?.sources || []).map((source) => [source.id, source]));
  const seen = new Set();
  const comparedPages = new Set();
  for (const page of value.visualReview.pageComparisons) {
    if (!page || typeof page !== "object" || Array.isArray(page)) throw new Error("visualReview.pageComparisons entries must be objects");
    const sourceId = String(page.sourceId || "");
    const pageNumber = page.page;
    const expected = expectedSources.get(sourceId);
    if (!REQUIRED_SOURCES.includes(sourceId) || !expected) throw new Error("visualReview.pageComparisons contains an unknown source");
    if (!Number.isSafeInteger(pageNumber) || pageNumber < 1 || pageNumber > expected.inventory.slideCount) {
      throw new Error(`${sourceId} page comparison has an invalid page number`);
    }
    const pageKey = `${sourceId}:${pageNumber}`;
    if (comparedPages.has(pageKey)) throw new Error(`duplicate page comparison for ${pageKey}`);
    comparedPages.add(pageKey);
    if (!/^[0-9a-f]{64}$/u.test(String(page.sourcePixelSha256 || "")) || !/^[0-9a-f]{64}$/u.test(String(page.outputPixelSha256 || ""))) {
      throw new Error(`${pageKey} must include source and output pixel SHA-256 values`);
    }
    const targetPage = expected.targets.some((target) => Number(String(target.nodeId).match(/^presentation\/slide\/(\d+)/u)?.[1]) === pageNumber);
    if (page.target !== targetPage) throw new Error(`${pageKey}.target does not match the frozen edit targets`);
    if (typeof page.pixelIdentical !== "boolean") throw new Error(`${pageKey}.pixelIdentical must be boolean`);
    if (targetPage) {
      if (page.pixelIdentical || page.sourcePixelSha256 === page.outputPixelSha256) throw new Error(`${pageKey} target page must show a pixel delta`);
    } else if (!page.pixelIdentical || page.sourcePixelSha256 !== page.outputPixelSha256) {
      throw new Error(`${pageKey} non-target page must have identical pixel hashes`);
    }
  }
  const expectedPageKeys = new Set(REQUIRED_SOURCES.flatMap((id) => {
    const source = expectedSources.get(id);
    return Array.from({ length: source?.inventory.slideCount || 0 }, (_, index) => `${id}:${index + 1}`);
  }));
  if (comparedPages.size !== expectedPageKeys.size || [...expectedPageKeys].some((key) => !comparedPages.has(key))) {
    throw new Error("visualReview.pageComparisons must cover every frozen sample page exactly once");
  }
  for (const id of REQUIRED_SOURCES) {
    const source = value.sources.find((candidate) => candidate?.id === id);
    if (!source) throw new Error(`missing evidence for ${id}`);
    if (seen.has(id)) throw new Error(`duplicate evidence for ${id}`);
    seen.add(id);
    if (!/^[0-9a-f]{64}$/u.test(String(source.sourceSha256 || ""))) throw new Error(`${id}.sourceSha256 must be a SHA-256`);
    const expected = expectedSources.get(id);
    if (expected && source.sourceSha256 !== expected.sha256) throw new Error(`${id}.sourceSha256 does not match the frozen benchmark manifest`);
    if (!/^[A-Za-z]:[\\/].+/u.test(String(source.sourcePath || "")) || !/^[A-Za-z]:[\\/].+/u.test(String(source.outputPath || ""))) {
      throw new Error(`${id} sourcePath and outputPath must be absolute Windows paths`);
    }
    if (source.outputPath === source.sourcePath) throw new Error(`${id} outputPath must differ from sourcePath`);
    if (!source.checks || typeof source.checks !== "object") throw new Error(`${id}.checks is required`);
    for (const check of REQUIRED_CHECKS) {
      if (source.checks[check] !== true) throw new Error(`${id}.checks.${check} must be true in human-observed evidence`);
    }
    if (!String(source.target?.nodeId || "").startsWith("presentation/slide/")) throw new Error(`${id}.target.nodeId must be a Presentation node locator`);
    if (!String(source.target?.operation || "")) throw new Error(`${id}.target.operation is required`);
    if (!String(source.evidencePath || "")) throw new Error(`${id}.evidencePath is required`);
  }

  return {
    schema: value.schema,
    checkedAt: value.checkedAt,
    commit: value.commit,
    platform: value.host.platform,
    powerpointVersion: value.host.powerpoint.version,
    sources: REQUIRED_SOURCES,
    pagesCompared: value.visualReview.pagesCompared,
  };
}

const entry = process.argv[1] ? path.resolve(process.argv[1]) : "";
if (entry === path.resolve(fileURLToPath(import.meta.url))) {
  const evidencePath = process.argv[2];
  const expectedCommit = process.argv[3];
  const manifestPath = process.argv[4] ? path.resolve(process.argv[4]) : DEFAULT_MANIFEST;
  if (!evidencePath) {
    console.error("usage: node scripts/validate-windows-pptx-lossless-evidence.mjs <evidence.json> [expected-commit] [manifest.json]");
    process.exit(2);
  }
  try {
    const [value, manifest] = await Promise.all([
      fs.readFile(path.resolve(evidencePath), "utf8").then(JSON.parse),
      fs.readFile(manifestPath, "utf8").then(JSON.parse),
    ]);
    console.log(JSON.stringify(validateWindowsPptxLosslessEvidence(value, { expectedCommit, manifest }), null, 2));
  } catch (error) {
    console.error(`Windows PPTX lossless evidence rejected: ${error.message}`);
    process.exit(1);
  }
}
