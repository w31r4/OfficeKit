#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const REPOSITORY_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DEFAULT_MANIFEST = path.join(REPOSITORY_ROOT, "evals/pptx-lossless/manifest.v1.json");
const REQUIRED_SOURCES = ["suanzhi-future-2026", "blue-gray-acid-template", "mckinsey-customer-loyalty"];
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
  if (value.visualReview.observedAt.slice(0, 10) !== value.checkedAt.slice(0, 10)) throw new Error("visualReview must be on the evidence date");
  if (value.visualReview.renderer !== "Microsoft PowerPoint") throw new Error("visualReview.renderer must be Microsoft PowerPoint");
  if (!Number.isInteger(value.visualReview.pagesCompared) || value.visualReview.pagesCompared < 1) throw new Error("visualReview.pagesCompared must be a positive integer");
  if (!String(value.visualReview.evidencePath || "")) throw new Error("visualReview.evidencePath is required");
  if (!/^(?:[0-9a-f]{40}|[0-9a-f]{64})$/u.test(String(value.commit || ""))) throw new Error("commit must be a 40-character SHA-1 or 64-character SHA-256");
  if (expectedCommit && value.commit !== expectedCommit) throw new Error("evidence commit does not match the checked-out commit");
  if (!Array.isArray(value.sources) || value.sources.length !== REQUIRED_SOURCES.length) throw new Error("evidence must contain exactly the three lossless PPTX sources");

  const expectedSources = new Map((manifest?.sources || []).map((source) => [source.id, source]));
  const seen = new Set();
  for (const id of REQUIRED_SOURCES) {
    const source = value.sources.find((candidate) => candidate?.id === id);
    if (!source) throw new Error(`missing evidence for ${id}`);
    if (seen.has(id)) throw new Error(`duplicate evidence for ${id}`);
    seen.add(id);
    if (!/^[0-9a-f]{64}$/u.test(String(source.sourceSha256 || ""))) throw new Error(`${id}.sourceSha256 must be a SHA-256`);
    const expected = expectedSources.get(id);
    if (expected && source.sourceSha256 !== expected.sha256) throw new Error(`${id}.sourceSha256 does not match the frozen benchmark manifest`);
    if (!String(source.sourcePath || "") || !String(source.outputPath || "")) throw new Error(`${id} sourcePath and outputPath are required`);
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
