import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const manifestPath = path.join(repoRoot, "evals", "presentation-authoring-compiler", "pilot.v1.json");
const manifest = JSON.parse(await fs.readFile(manifestPath, "utf8"));

assert.equal(manifest.schema, "office-kit/presentation-authoring-pilot/v1");
assert.equal(manifest.status, "machine-complete-blind-judged-default-kept");
assert.equal(manifest.baseline.arm, "A");
assert.equal(manifest.baseline.route, "grid-default");
assert.equal(manifest.baseline.frozenAt, "origin/main@a0452867");
assert.equal(manifest.baseline.defaultUntilThresholdsPass, true);
assert.deepEqual(Object.keys(manifest.arms).sort(), ["A", "B", "C"]);
assert.equal(manifest.tasks.length, 10);
assert.equal(new Set(manifest.tasks.map((task) => task.scenario)).size, 5);
assert.equal(manifest.design.trialsPerTask * manifest.design.armsPerTask * manifest.tasks.length, manifest.design.totalRuns);
assert.equal(manifest.design.sameInputsAcrossArms, true);
assert.equal(manifest.design.freshContextPerRun, true);
assert.equal(manifest.design.package, "packed clean-install");
for (const threshold of Object.values(manifest.thresholds)) assert.ok([">=", "<=", ">"].includes(threshold.operator));
assert.match(manifest.rolloutRule, /every threshold passes.*otherwise keep A shipped.*C.*experimental/i);
const resultsPath = path.join(repoRoot, "evals", "presentation-authoring-compiler", "results.v1.json");
const results = JSON.parse(await fs.readFile(resultsPath, "utf8"));
assert.equal(results.observedRuns, manifest.design.totalRuns);
assert.equal(results.rollout.status, "keep-A");
assert.equal(results.blind.judgments, 40);
assert.equal(results.blind.overA, 0.5);
assert.equal(results.blind.overB, 0.6);
assert.equal(results.thresholds.blindWinRateOverA.status, "failed");
assert.equal(results.thresholds.blindWinRateOverB.status, "passed");
assert.equal(results.evidence.blindPacketStatus, "judged");
assert.equal(results.evidence.blindJudgmentCount, 40);
assert.equal(results.evidence.blindReview, "evals/presentation-authoring-compiler/blind-review.v1.json");
const blindReview = JSON.parse(await fs.readFile(path.join(repoRoot, results.evidence.blindReview), "utf8"));
assert.equal(blindReview.status, "complete");
assert.equal(blindReview.reviewer.freshContextPerComparison, true);
assert.equal(blindReview.reviewer.sourceCodeAccess, false);
assert.equal(blindReview.comparisons, 20);
assert.equal(blindReview.pairwiseJudgments, 40);
assert.equal(blindReview.summary.cOverA, 0.5);
assert.equal(blindReview.summary.cOverB, 0.6);
const judgmentLines = (await fs.readFile(path.join(repoRoot, "evals/presentation-authoring-compiler/judgments.v1.jsonl"), "utf8"))
  .split(/\r?\n/u)
  .filter(Boolean)
  .map((line) => JSON.parse(line));
assert.equal(judgmentLines.length, 40);
assert.equal(new Set(judgmentLines.map((entry) => entry.comparisonKey)).size, 40);
assert.ok(judgmentLines.every((entry) => entry.leftArm === "C" && ["A", "B"].includes(entry.rightArm)));
assert.ok(judgmentLines.every((entry) => ["left", "right", "tie"].includes(entry.winner)));
assert.doesNotMatch(JSON.stringify(manifest), /(?:\/Users\/|[A-Z]:\\|\/tmp\/)/u);

console.log("presentation authoring evaluation manifest ok");
