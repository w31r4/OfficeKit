import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

import { buildPilotMatrix, buildPilotPrompt, loadPilotManifest } from "../scripts/presentation-authoring-pilot.mjs";
import { scorePilot } from "../scripts/score-presentation-authoring-pilot.mjs";

const manifest = await loadPilotManifest();
const matrix = buildPilotMatrix(manifest);
assert.equal(matrix.length, 60);
assert.equal(buildPilotMatrix(manifest, { taskId: "business-review-01", arm: "C", trial: 1 }).length, 1);
assert.match(buildPilotPrompt({ manifest, ...buildPilotMatrix(manifest, { taskId: "business-review-01", arm: "C", trial: 1 })[0] }), /independent final review[\s\S]*without a source baseline[\s\S]*Do not read \.office-kit\/tasks/i);
assert.match(buildPilotPrompt({ manifest, ...buildPilotMatrix(manifest, { taskId: "business-review-01", arm: "C", trial: 1 })[0] }), /minimumBodyFontSize[\s\S]*minimumCaptionFontSize[\s\S]*18[\s\S]*four distinct composition silhouettes[\s\S]*cardWallPattern[\s\S]*no unrecorded design warnings[\s\S]*intentionalWarnings[\s\S]*dominant reading anchor[\s\S]*contrast[\s\S]*pale card surface[\s\S]*quantitative claim/i);
assert.match(buildPilotPrompt({ manifest, ...buildPilotMatrix(manifest, { taskId: "business-review-01", arm: "C", trial: 1 })[0] }), /private store[\s\S]*never use shell or node[\s\S]*evidence paths/i);
assert.doesNotMatch(buildPilotPrompt({ manifest, ...buildPilotMatrix(manifest, { taskId: "business-review-01", arm: "A", trial: 1 })[0] }), /minimumBodyFontSize/u);
const pilotSource = await readFile(new URL("../scripts/presentation-authoring-pilot.mjs", import.meta.url), "utf8");
assert.match(pilotSource, /unresolved-design-warnings/u);
assert.match(pilotSource, /authoringPlanPath/u);
assert.equal(new Set(matrix.map((entry) => entry.task.id)).size, 10);
assert.deepEqual(new Set(matrix.map((entry) => entry.arm)), new Set(["A", "B", "C"]));
assert.equal(new Set(matrix.map((entry) => entry.armOrder.join(""))).size > 1, true, "arm order must be deterministically randomized");

const incomplete = scorePilot(manifest, []);
assert.equal(incomplete.rollout.status, "keep-A");
assert.equal(incomplete.thresholds.hardPassRate.status, "insufficient-evidence");
assert.equal(incomplete.thresholds.blindWinRateOverA.status, "insufficient-evidence");

const passingRuns = matrix.map(({ task, arm, trial, armOrder }) => ({
  schema: "office-kit/presentation-authoring-pilot-run/v1",
  runId: `${task.id}/${arm}/${trial}`,
  taskId: task.id,
  scenario: task.scenario,
  arm,
  trial,
  armOrder,
  elapsedMs: arm === "A" ? 100 : 110,
  tokenUsage: { observed: true, totalTokens: 1_000 },
  status: "passed",
  checks: { task: { passed: true } },
}));
const judgments = [];
for (const { task, trial } of matrix.filter(({ arm }) => arm === "A")) {
  judgments.push({ comparisonKey: `${task.id}:${trial}`, leftArm: "C", rightArm: "A", winner: "left" });
  judgments.push({ comparisonKey: `${task.id}:${trial}`, leftArm: "C", rightArm: "B", winner: "left" });
}
const passing = scorePilot(manifest, passingRuns, judgments);
assert.equal(passing.rollout.status, "switch-C");
assert.equal(passing.thresholds.hardPassRate.status, "passed");
assert.equal(passing.thresholds.blindWinRateOverA.status, "passed");
assert.equal(passing.thresholds.medianTimeAndTokensRatioToA.status, "passed");

console.log("presentation authoring pilot contract ok");
