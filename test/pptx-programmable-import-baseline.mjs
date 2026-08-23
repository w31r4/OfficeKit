import assert from "node:assert/strict";

import { buildBaselineEvidence, buildCandidateEvidence } from "../scripts/pptx-programmable-import-baseline.mjs";
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

const candidateMatrix = structuredClone(matrix);
for (const source of candidateMatrix.sources) {
  source.sourceSha256 = "1".repeat(64);
  for (const intent of source.intents) {
    intent.deterministic = true;
    for (const run of intent.runs) {
      run.status = "passed";
      run.sourceSha256After = source.sourceSha256;
      run.outputSha256 = "2".repeat(64);
      run.worker = { sourceUnchanged: true, secondImport: true, outputSha256: run.outputSha256 };
      run.packageOracle = {
        partSet: { passed: true },
        nonTargetPartsByteIdentical: true,
        relationships: { passed: true },
        targetMask: { passed: true },
      };
      run.pixelOracle = { passed: true, targetPageChanged: true, nonTargetPagesPixelIdentical: true, nonTargetMismatches: [] };
    }
  }
}
candidateMatrix.acceptance = { requiredIntents: 30, requiredRuns: 90, passedRuns: 90, deterministicIntents: 30, status: "passed" };

const candidateCodex = structuredClone(codex);
for (const trial of candidateCodex.trials) {
  trial.status = "passed";
  trial.failures = [];
  const outputSha256 = `${Math.floor(candidateCodex.trials.indexOf(trial) / 3) + 7}`.repeat(64);
  trial.checks = {
    codex: { passed: true },
    policy: { passed: true, findings: [] },
    source: { passed: true },
    output: { passed: true, createOnly: true, outputCount: 1 },
    durableTask: {
      passed: true,
      sessions: 3,
      commits: [{ commitId: "c0001" }, { commitId: "c0002" }],
      head: { commitId: "c0002", revisionSha256: outputSha256 },
      pending: [],
      publication: { commitId: "c0002", sha256: outputSha256 },
    },
    packageOracle: { outputSha256, partSet: { passed: true }, relationships: { passed: true }, targetMask: { passed: true } },
    secondImport: { passed: true, inputSha256: outputSha256 },
    pixelOracle: { passed: true, appendedTargetChangedFromSource: true, nonTargetPagesPixelIdentical: true, nonTargetMismatches: [] },
  };
}
candidateCodex.acceptance = { requiredTasks: 3, trialsPerTask: 3, requiredTrials: 9, completedTrials: 9, passedTrials: 9, status: "passed" };

const companion = {
  schema: "office-kit/pptx-source-derived-companion-evidence/v1",
  productBaseline: matrix.baseline,
  definitionsSha256: "3".repeat(64),
  repetitionsPerCase: 3,
  package: { name: "office-kit", version: "0.6.0", installKind: "packed-clean-install", tarballSha256: "4".repeat(64) },
  environment: { node: "synthetic" },
  acceptance: { scope: "full-suite", status: "passed" },
  coverage: {
    required: ["text", "geometry", "image", "table", "chart", "component", "add", "delete", "reorder"],
    passed: ["text", "geometry", "image", "table", "chart", "component", "add", "delete", "reorder"],
    missing: [],
    status: "passed",
  },
  cases: [["text", "add"], ["geometry"], ["image"], ["table"], ["chart"], ["component"], ["delete"], ["reorder"]].map((covers, caseIndex) => ({
    id: `synthetic-companion-${caseIndex + 1}`,
    covers,
    requiredRuns: 3,
    completedRuns: 3,
    passedRuns: 3,
    deterministic: true,
    runs: Array.from({ length: 3 }, (_, index) => ({
      repetition: index + 1,
      status: "passed",
      outputSha256: String(caseIndex + 1).repeat(64),
      worker: { sourceUnchanged: true, secondImport: { passed: true } },
      packageOracle: { passed: true, partSet: { passed: true }, nonTargetPartsByteIdentical: true, targetMask: { passed: true } },
      pixelOracle: { passed: true, targetPageChanged: true, nonTargetPagesPixelIdentical: true, nonTargetMismatches: [] },
    })),
  })),
};
const candidateMatrixBytes = Buffer.from(`${JSON.stringify(candidateMatrix)}\n`);
const candidateCodexBytes = Buffer.from(`${JSON.stringify(candidateCodex)}\n`);
const companionBytes = Buffer.from(`${JSON.stringify(companion)}\n`);
const candidate = buildCandidateEvidence({
  matrix: candidateMatrix,
  codex: candidateCodex,
  companion,
  matrixBytes: candidateMatrixBytes,
  codexBytes: candidateCodexBytes,
  companionBytes,
  candidateHead: "6".repeat(40),
});
assert.equal(candidate.schema, "office-kit/pptx-programmable-import-candidate/v1");
assert.equal(candidate.acceptance.status, "passed");
assert.equal(candidate.matrix.passedRuns, 90);
assert.equal(candidate.codex.passedTrials, 9);
assert.equal(candidate.sourceDerived.passedRuns, 24);
assert.equal(candidate.evidenceFiles.companion.sha256, sha256(companionBytes));
assert.throws(
  () => buildCandidateEvidence({
    matrix: { ...candidateMatrix, package: { ...candidateMatrix.package, tarballSha256: "7".repeat(64) } },
    codex: candidateCodex,
    companion,
    matrixBytes: candidateMatrixBytes,
    codexBytes: candidateCodexBytes,
    companionBytes,
    candidateHead: "6".repeat(40),
  }),
  /same deterministic tarball/u,
);
const matrixWithFailedOracle = structuredClone(candidateMatrix);
matrixWithFailedOracle.sources[0].intents[0].runs[0].packageOracle.targetMask.passed = false;
assert.throws(
  () => buildCandidateEvidence({ matrix: matrixWithFailedOracle, codex: candidateCodex, companion, matrixBytes: candidateMatrixBytes, codexBytes: candidateCodexBytes, companionBytes, candidateHead: "6".repeat(40) }),
  /package oracle did not pass/u,
);
const codexWithoutDurableTask = structuredClone(candidateCodex);
codexWithoutDurableTask.trials[0].checks.durableTask.passed = false;
assert.throws(
  () => buildCandidateEvidence({ matrix: candidateMatrix, codex: codexWithoutDurableTask, companion, matrixBytes: candidateMatrixBytes, codexBytes: candidateCodexBytes, companionBytes, candidateHead: "6".repeat(40) }),
  /Codex oracle set did not pass/u,
);
const incompleteCompanion = structuredClone(companion);
incompleteCompanion.coverage.passed = incompleteCompanion.coverage.passed.filter((kind) => kind !== "reorder");
assert.throws(
  () => buildCandidateEvidence({ matrix: candidateMatrix, codex: candidateCodex, companion: incompleteCompanion, matrixBytes: candidateMatrixBytes, codexBytes: candidateCodexBytes, companionBytes, candidateHead: "6".repeat(40) }),
  /every required operation category/u,
);

console.log("PPTX programmable-import baseline evidence builder ok");
