import assert from "node:assert/strict";

import { buildBaselineEvidence } from "../scripts/pptx-programmable-import-baseline.mjs";
import { sha256 } from "../scripts/pptx-programmable-import-oracle.mjs";

const tarballSha256 = "a".repeat(64);
const matrix = {
  schema: "office-kit/pptx-programmable-import-matrix/v1",
  baseline: "d5df8df94727dccd4412e6be874d1c5407b57f64",
  definitionsSha256: "b".repeat(64),
  repetitionsPerIntent: 3,
  package: { name: "office-kit", version: "0.6.0", installKind: "packed-clean-install", tarballSha256 },
  environment: { node: "synthetic" },
  sources: Array.from({ length: 3 }, (_, sourceIndex) => ({
    id: `source-${sourceIndex + 1}`,
    intents: Array.from({ length: 10 }, (_, intentIndex) => {
      const index = sourceIndex * 10 + intentIndex;
      return {
        id: `intent-${index + 1}`,
        requiredRuns: 3,
        completedRuns: 3,
        deterministic: index !== 0,
        passedRuns: 3,
        runs: Array.from({ length: 3 }, (__, repetition) => ({ repetition: repetition + 1, status: "passed", outputSha256: index === 0 ? String(repetition).repeat(64) : "c".repeat(64) })),
      };
    }),
  })),
  acceptance: { requiredIntents: 30, requiredRuns: 90, passedRuns: 90, deterministicIntents: 29, status: "failed" },
};
const codex = {
  schema: "office-kit/pptx-codex-continuation-evidence/v1",
  baseline: matrix.baseline,
  definitions: { continuationSha256: "d".repeat(64), intentSha256: matrix.definitionsSha256 },
  package: { name: "office-kit", version: "0.6.0", tarballSha256, cleanInstallPerTrial: true },
  environment: { codex: "synthetic" },
  trials: Array.from({ length: 9 }, (_, index) => ({
    taskId: `task-${Math.floor(index / 3) + 1}`,
    sourceId: `source-${Math.floor(index / 3) + 1}`,
    repetition: (index % 3) + 1,
    status: index === 0 ? "failed" : "passed",
    failures: index === 0 ? ["preserved failure"] : [],
    freshCodexContext: true,
    packedCleanInstall: { passed: true, package: "office-kit", version: "0.6.0", tarballSha256 },
    checks: { codex: { status: 0, agentFinal: index === 0 ? "真实失败原因" : "完成" }, policy: { passed: true }, output: { passed: index !== 0 } },
  })),
  acceptance: { requiredTasks: 3, trialsPerTask: 3, requiredTrials: 9, completedTrials: 9, passedTrials: 8, status: "failed" },
};
const matrixBytes = Buffer.from(`${JSON.stringify(matrix)}\n`);
const codexBytes = Buffer.from(`${JSON.stringify(codex)}\n`);
const baseline = buildBaselineEvidence({ matrix, codex, matrixBytes, codexBytes, harnessHead: "e".repeat(40) });
assert.equal(baseline.acceptance.status, "failed");
assert.equal(baseline.acceptance.failuresPreserved, true);
assert.equal(baseline.matrix.nonDeterministic.length, 1);
assert.equal(baseline.matrix.nonDeterministic[0].outputSha256s.length, 3);
assert.equal(baseline.codex.failedTrials.length, 1);
assert.equal(baseline.codex.failedTrials[0].agentFinal, "真实失败原因");
assert.equal(baseline.evidenceFiles.matrix.sha256, sha256(matrixBytes));
assert.equal(baseline.evidenceFiles.codex.sha256, sha256(codexBytes));
assert.throws(
  () => buildBaselineEvidence({ matrix: { ...matrix, package: { ...matrix.package, tarballSha256: "f".repeat(64) } }, codex, matrixBytes, codexBytes, harnessHead: "e".repeat(40) }),
  /same deterministic tarball/u,
);
assert.throws(
  () => buildBaselineEvidence({ matrix, codex: { ...codex, trials: codex.trials.slice(0, 8) }, matrixBytes, codexBytes, harnessHead: "e".repeat(40) }),
  /nine trials/u,
);
assert.throws(
  () => buildBaselineEvidence({ matrix: { ...matrix, sources: matrix.sources.slice(0, 2) }, codex, matrixBytes, codexBytes, harnessHead: "e".repeat(40) }),
  /three distinct sources/u,
);

console.log("PPTX programmable-import baseline evidence builder ok");
