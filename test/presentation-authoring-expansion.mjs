import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const expansionPath = path.join(repoRoot, "evals/presentation-authoring-compiler/expansion.v1.json");
const expansion = JSON.parse(await readFile(expansionPath, "utf8"));
assert.equal(expansion.schema, "office-kit/presentation-authoring-expansion/v1");
assert.equal(expansion.route, "C");
assert.equal(expansion.sourcePolicy, "brief-only");
assert.equal(expansion.taskCount, 30);
assert.equal(expansion.scenarioCount, 5);
assert.equal(expansion.tasksPerScenario, 6);
assert.equal(expansion.tasks.length, 30);
assert.equal(new Set(expansion.tasks.map((task) => task.id)).size, 30);
assert.deepEqual([...new Set(expansion.tasks.map((task) => task.scenario))].sort(), [
  "academic-defense",
  "analytical-reporting",
  "brand-launch",
  "business-review",
  "technical-architecture",
]);
for (const task of expansion.tasks) {
  assert.match(task.id, /^[a-z0-9]+(?:-[a-z0-9]+)+$/u);
  for (const field of ["goal", "contentShape", "revisionIntent"]) assert.ok(task[field]?.trim(), `${task.id}: ${field}`);
}
assert.equal(expansion.execution.runner, "scripts/presentation-authoring-expansion.mjs");
assert.match(await readFile(path.join(repoRoot, expansion.execution.runner), "utf8"), /runPilotTrial/u);
assert.equal(expansion.execution.fullMatrixRequired, true);
if (expansion.execution.status === "registered") {
  assert.equal(expansion.execution.fullMatrixRuns, 0);
} else {
  assert.equal(expansion.execution.status, "completed");
  assert.equal(expansion.execution.fullMatrixRuns, expansion.taskCount);
  assert.equal(expansion.execution.fullMatrixPassed, expansion.taskCount);
  assert.equal(expansion.execution.evidence, "evals/presentation-authoring-compiler/expansion-runs.v1.json");
  const fullMatrix = JSON.parse(await readFile(path.join(repoRoot, expansion.execution.evidence), "utf8"));
  assert.equal(fullMatrix.acceptance.expectedRuns, expansion.taskCount);
  assert.equal(fullMatrix.acceptance.completedRuns, expansion.taskCount);
  assert.equal(fullMatrix.acceptance.passedRuns, expansion.taskCount);
  assert.equal(fullMatrix.acceptance.status, "passed");
  const { execution: _execution, ...expansionDefinition } = expansion;
  const expectedDefinitionHash = (await import("node:crypto")).createHash("sha256").update(JSON.stringify(expansionDefinition)).digest("hex");
  assert.equal(fullMatrix.expansionDefinitionSha256, expectedDefinitionHash);
  assert.doesNotMatch(JSON.stringify(fullMatrix), /(?:\/Users\/|[A-Z]:\\|\/tmp\/)/u);
}
const continuation = expansion.selectedContinuation;
assert.equal(continuation.completed, 23);
assert.equal(continuation.passed, 23);
assert.equal(continuation.successRate, 1);
assert.ok(continuation.successRate >= continuation.requiredSuccessRate);
assert.equal(continuation.status, "passed");
assert.deepEqual(continuation.evidence.map((entry) => entry.path), [
  "evals/presentation-authoring-compiler/postfix-c.v7.json",
  "evals/presentation-authoring-compiler/postfix-c.v8.json",
]);
assert.doesNotMatch(JSON.stringify(expansion), /(?:\/Users\/|[A-Z]:\\|\/tmp\/)/u);

const pilot = JSON.parse(await readFile(path.join(repoRoot, "evals/presentation-authoring-compiler/pilot.v1.json"), "utf8"));
assert.deepEqual(pilot.tasks.map((task) => task.id).filter((id) => expansion.tasks.some((candidate) => candidate.id === id)), []);
const v7 = JSON.parse(await readFile(path.join(repoRoot, "evals/presentation-authoring-compiler/postfix-c.v7.json"), "utf8"));
const selectedIds = new Set(v7.mainRuns.map((run) => run.taskId));
assert.equal(selectedIds.size, 10);
assert.ok(v7.mainRuns.every((run) => run.arm === "C" && run.status === "passed" && run.plan === true && run.commits >= 1 && run.publications >= 1));
const v8 = JSON.parse(await readFile(path.join(repoRoot, "evals/presentation-authoring-compiler/postfix-c.v8.json"), "utf8"));
assert.equal(v8.unseenHoldout.runs.C.completed, v8.unseenHoldout.runs.C.passed);

console.log("presentation authoring expansion contract ok (30 tasks registered, continuation 23/23)");
