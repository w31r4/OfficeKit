#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";

const MAX_MANIFEST_BYTES = 256 * 1024;
const RESTORED_THRESHOLD = 95;

const VISUAL_DIMENSIONS = Object.freeze([
  ["silhouette", 16],
  ["hierarchy", 14],
  ["paletteAndSurfaces", 12],
  ["typography", 10],
  ["densityAndRhythm", 10],
  ["visualCarriers", 12],
  ["layerRelationships", 12],
  ["motifs", 8],
  ["exampleCoverage", 6],
]);

const FUNCTIONAL_DIMENSIONS = Object.freeze([
  ["inspectDiscovery", 14],
  ["editableLeaves", 14],
  ["reusableAssets", 12],
  ["roundTripStability", 16],
  ["nativeRendering", 14],
  ["backgroundAndLayerFidelity", 12],
  ["opaquePreservation", 10],
  ["safeRefusal", 8],
]);

let args;
try {
  args = parseArgs(process.argv.slice(2));
  const report = await scoreManifest(args.manifestPath);
  process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
  if (!report.restored) process.exitCode = 1;
} catch (error) {
  const message = error instanceof Error ? error.message : String(error);
  process.stderr.write(`${JSON.stringify({ ok: false, error: message })}\n`);
  process.exitCode = 2;
}

async function scoreManifest(manifestPath) {
  const absolutePath = path.resolve(manifestPath);
  const stat = await fs.stat(absolutePath);
  if (!stat.isFile() || stat.size > MAX_MANIFEST_BYTES) {
    throw new Error(`fidelity manifest must be a regular file of at most ${MAX_MANIFEST_BYTES} bytes`);
  }
  let manifest;
  try {
    manifest = JSON.parse(await fs.readFile(absolutePath, "utf8"));
  } catch (error) {
    throw new Error(`fidelity manifest is not valid JSON: ${error.message}`);
  }
  if (manifest == null || typeof manifest !== "object" || Array.isArray(manifest)) {
    throw new Error("fidelity manifest must be an object");
  }
  if (manifest.schemaVersion !== 1) throw new Error("fidelity manifest schemaVersion must be 1");
  if (typeof manifest.templateId !== "string" || manifest.templateId.trim() === "") {
    throw new Error("fidelity manifest templateId must be a non-empty string");
  }
  validateEvidenceBundle(manifest.evidence);
  const visual = scoreSection(manifest.visual, VISUAL_DIMENSIONS, "visual");
  const functional = scoreSection(manifest.functional, FUNCTIONAL_DIMENSIONS, "functional");
  return {
    schemaVersion: 1,
    templateId: manifest.templateId,
    visual,
    functional,
    threshold: RESTORED_THRESHOLD,
    restored: visual.score >= RESTORED_THRESHOLD && functional.score >= RESTORED_THRESHOLD,
  };
}

function scoreSection(section, dimensions, label) {
  if (section == null || typeof section !== "object" || Array.isArray(section)) {
    throw new Error(`${label} evidence must be an object`);
  }
  const expected = new Set(dimensions.map(([id]) => id));
  const actual = new Set(Object.keys(section));
  const missing = dimensions.map(([id]) => id).filter((id) => !actual.has(id));
  const extra = [...actual].filter((id) => !expected.has(id));
  if (missing.length > 0) throw new Error(`${label} evidence is missing: ${missing.join(", ")}`);
  if (extra.length > 0) throw new Error(`${label} evidence contains unsupported dimensions: ${extra.join(", ")}`);

  let weighted = 0;
  let totalWeight = 0;
  const scored = {};
  for (const [id, weight] of dimensions) {
    const entry = section[id];
    if (entry == null || typeof entry !== "object" || Array.isArray(entry)) {
      throw new Error(`${label}.${id} must contain score and evidence`);
    }
    const score = Number(entry.score);
    if (!Number.isFinite(score) || score < 0 || score > 100) {
      throw new Error(`${label}.${id}.score must be a finite number from 0 to 100`);
    }
    if (!Array.isArray(entry.evidence) || entry.evidence.length === 0 ||
        entry.evidence.some((value) => typeof value !== "string" || value.trim() === "")) {
      throw new Error(`${label}.${id}.evidence must contain at least one non-empty path or locator`);
    }
    const roundedScore = Math.round(score * 100) / 100;
    weighted += roundedScore * weight;
    totalWeight += weight;
    scored[id] = { score: roundedScore, evidence: [...entry.evidence] };
  }
  return {
    score: Math.round((weighted / totalWeight) * 100) / 100,
    weights: Object.fromEntries(dimensions),
    dimensions: scored,
  };
}

function validateEvidenceBundle(evidence) {
  const required = ["source", "renders", "inspect", "edits", "reimport", "package"];
  if (evidence == null || typeof evidence !== "object" || Array.isArray(evidence)) {
    throw new Error(`evidence must contain ${required.join(", ")}`);
  }
  const extras = Object.keys(evidence).filter((key) => !required.includes(key));
  if (extras.length > 0) throw new Error(`evidence contains unsupported fields: ${extras.join(", ")}`);
  for (const key of required) {
    const values = Array.isArray(evidence[key]) ? evidence[key] : [evidence[key]];
    if (values.length === 0 || values.some((value) => typeof value !== "string" || value.trim() === "")) {
      throw new Error(`evidence.${key} must contain at least one path or locator`);
    }
  }
}

function parseArgs(argv) {
  if (argv.length !== 2 || argv[0] !== "--manifest" || argv[1].startsWith("--")) {
    throw new Error("Usage: presentation-template-fidelity.mjs --manifest <evidence.json>");
  }
  return { manifestPath: argv[1] };
}
