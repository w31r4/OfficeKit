import assert from "node:assert/strict";
import { mkdtemp, readFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";

const repoRoot = path.resolve(import.meta.dirname, "..");
const assetsDir = process.env.OFFICEKIT_PPTX_ACCEPTANCE_ASSETS || "/Users/zfang/Downloads/飞书20260814-175228";
const runParent = await mkdtemp(path.join(os.tmpdir(), "officekit-pptx-programmable-matrix-test-"));
const runRoot = path.join(runParent, "matrix");
const result = spawnSync(process.execPath, [
  "scripts/pptx-programmable-import-matrix.mjs",
  "--assets-dir", assetsDir,
  "--package-root", repoRoot,
  "--run-root", runRoot,
  "--source", "blue-gray-acid-template",
  "--intent", "cover-subtitle",
  "--repetitions", "2",
  "--no-render",
], { cwd: repoRoot, encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
const evidence = JSON.parse(await readFile(path.join(runRoot, "evidence.json"), "utf8"));
assert.equal(evidence.acceptance.requiredIntents, 1);
assert.equal(evidence.acceptance.requiredRuns, 2);
assert.equal(evidence.acceptance.passedRuns, 2);
assert.equal(evidence.acceptance.deterministicIntents, 1);
assert.equal(evidence.acceptance.status, "passed");
const runs = evidence.sources[0].intents[0].runs;
assert.equal(new Set(runs.map(({ outputSha256 }) => outputSha256)).size, 1);
assert.ok(runs.every(({ worker, packageOracle, pixelOracle }) => worker.secondImport === true && packageOracle.nonTargetPartsByteIdentical === true && packageOracle.relationships.passed === true && pixelOracle.skipped === true));

const svgRunRoot = path.join(runParent, "svg-nondeterminism");
const svg = spawnSync(process.execPath, [
  "scripts/pptx-programmable-import-matrix.mjs",
  "--assets-dir", assetsDir,
  "--package-root", repoRoot,
  "--run-root", svgRunRoot,
  "--source", "mckinsey-customer-loyalty",
  "--intent", "cover-programme-spelling",
  "--repetitions", "2",
  "--no-render",
], { cwd: repoRoot, encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
assert.equal(svg.status, 1, `${svg.stdout}\n${svg.stderr}`);
const svgEvidence = JSON.parse(await readFile(path.join(svgRunRoot, "evidence.json"), "utf8"));
assert.equal(svgEvidence.acceptance.passedRuns, 2, "both SVG edits should pass the independent per-run oracle");
assert.equal(svgEvidence.acceptance.deterministicIntents, 0, "random copy-on-write relationship IDs must not be normalized into a deterministic pass");
assert.equal(svgEvidence.acceptance.status, "failed");
assert.equal(new Set(svgEvidence.sources[0].intents[0].runs.map(({ outputSha256 }) => outputSha256)).size, 2);

const collision = spawnSync(process.execPath, [
  "scripts/pptx-programmable-import-matrix.mjs",
  "--assets-dir", assetsDir,
  "--run-root", runRoot,
  "--source", "blue-gray-acid-template",
  "--intent", "cover-subtitle",
  "--no-render",
], { cwd: repoRoot, encoding: "utf8" });
assert.notEqual(collision.status, 0);
assert.match(collision.stderr, /outputs are create-only/u);

console.log("PPTX programmable-import matrix smoke ok");
