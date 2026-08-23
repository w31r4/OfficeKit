import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

const BASELINE = "d5df8df94727dccd4412e6be874d1c5407b57f64";
const evidencePath = path.resolve(import.meta.dirname, "../evals/pptx-programmable-import/baseline/codex.v1.json");
const bytes = await readFile(evidencePath);
const text = bytes.toString("utf8");
const evidence = JSON.parse(text);

assert.equal(evidence.schema, "office-kit/pptx-codex-continuation-evidence/v1");
assert.equal(evidence.baseline, BASELINE);
assert.equal(evidence.package.name, "office-kit");
assert.equal(evidence.package.version, "0.6.0");
assert.equal(evidence.package.baselineCandidate, true);
assert.equal(evidence.package.cleanInstallPerTrial, true);
assert.equal(evidence.package.lifecycleScripts, "ignored");
assert.match(evidence.package.tarballSha256, /^[0-9a-f]{64}$/u);
assert.ok(evidence.package.packedBytes > 0);
assert.ok(evidence.package.unpackedBytes > evidence.package.packedBytes);
assert.ok(evidence.package.totalFiles > 0);
assert.deepEqual(evidence.protocol, {
  freshCodexContextPerTrial: true,
  replSessionsPerTrial: 3,
  publicPackageOnly: true,
  createOnlyOutputs: true,
});
assert.equal(evidence.environment.render, true);

assert.equal(evidence.trials.length, 9);
const taskGroups = Map.groupBy(evidence.trials, ({ taskId }) => taskId);
assert.equal(taskGroups.size, 3);
for (const [taskId, trials] of taskGroups) {
  assert.deepEqual(trials.map(({ repetition }) => repetition).sort(), [1, 2, 3], `${taskId}: exact repetitions`);
  assert.equal(new Set(trials.map(({ sourceId }) => sourceId)).size, 1, `${taskId}: stable source`);
}

for (const trial of evidence.trials) {
  const label = `${trial.taskId}/${trial.repetition}`;
  assert.equal(trial.freshCodexContext, true, `${label}: fresh Codex context`);
  assert.deepEqual(trial.packedCleanInstall, {
    passed: true,
    package: "office-kit",
    version: evidence.package.version,
    tarballSha256: evidence.package.tarballSha256,
  }, `${label}: isolated packed install`);
  assert.match(trial.evidenceDirectory, /^trials\/[^/]+\/[123]\/evaluator$/u, `${label}: relative evidence path`);
  assert.equal(trial.checks.source.passed, true, `${label}: immutable source`);
  assert.equal(trial.checks.source.expectedSha256, trial.checks.source.afterSha256, `${label}: source hash preserved`);
  assert.equal(trial.checks.source.immutableMode, 0o444, `${label}: read-only source mode`);
  assert.equal(trial.checks.policy.passed, true, `${label}: public API policy`);
  assert.deepEqual(trial.checks.policy.findings, [], `${label}: no forbidden implementation path`);
  assert.equal(trial.checks.codex.status, 0, `${label}: Codex exited deliberately`);
  assert.equal(trial.checks.codex.timedOut, false, `${label}: no timeout`);
  assert.equal(trial.checks.codex.captureExceeded, false, `${label}: complete bounded trace capture`);
  assert.equal(trial.checks.codex.agentFinalTruncated, false, `${label}: complete final reason`);
  assert.ok(trial.checks.codex.agentFinal?.trim(), `${label}: final reason retained`);
  assert.equal(trial.status, "failed", `${label}: frozen baseline status`);
  assert.ok(trial.failures.length > 0, `${label}: machine failure retained`);
  assert.equal(trial.checks.output.passed, false, `${label}: no output falsely accepted`);
}

assert.deepEqual(evidence.acceptance, {
  requiredTasks: 3,
  trialsPerTask: 3,
  requiredTrials: 9,
  completedTrials: 9,
  passedTrials: 0,
  status: "failed",
});
assert.doesNotMatch(text, /(?:\/Users\/|\/tmp\/)/u, "committed evidence must not retain machine-local absolute paths");

console.log("PPTX Codex continuation evidence ok (0/9 baseline failures preserved)");
