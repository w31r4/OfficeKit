import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

import { buildBaselineEvidence } from "../scripts/pptx-programmable-import-baseline.mjs";
import { sha256 } from "../scripts/pptx-programmable-import-oracle.mjs";

const BASELINE = "d5df8df94727dccd4412e6be874d1c5407b57f64";
const root = path.resolve(import.meta.dirname, "../evals/pptx-programmable-import");
const [matrixBytes, codexBytes, baselineBytes, definitionBytes] = await Promise.all([
  readFile(path.join(root, "baseline/matrix.v1.json")),
  readFile(path.join(root, "baseline/codex.v1.json")),
  readFile(path.join(root, "baseline.v1.json")),
  readFile(path.join(root, "intent-matrix.v1.json")),
]);
const matrix = JSON.parse(matrixBytes);
const codex = JSON.parse(codexBytes);
const baseline = JSON.parse(baselineBytes);
const definitions = JSON.parse(definitionBytes);

assert.equal(matrix.schema, "office-kit/pptx-programmable-import-matrix/v1");
assert.equal(matrix.baseline, BASELINE);
assert.equal(matrix.definitionsSha256, sha256(definitionBytes));
assert.equal(matrix.package.name, "office-kit");
assert.equal(matrix.package.version, "0.6.0");
assert.equal(matrix.package.installKind, "packed-clean-install");
assert.equal(matrix.package.tarballSha256, codex.package.tarballSha256);
assert.equal(matrix.environment.render, true);
assert.equal(matrix.repetitionsPerIntent, 3);
assert.equal(matrix.sources.length, 3);

let passedRuns = 0;
let deterministicIntents = 0;
for (const source of matrix.sources) {
  const declaredSource = definitions.sources.find(({ id }) => id === source.id);
  assert.ok(declaredSource, `${source.id}: declared source`);
  assert.equal(source.sourceSha256, declaredSource.sha256, `${source.id}: pinned source hash`);
  assert.equal(source.intents.length, 10, `${source.id}: ten intents`);
  for (const intent of source.intents) {
    const declaredIntent = declaredSource.intents.find(({ id }) => id === intent.id);
    const label = `${source.id}/${intent.id}`;
    assert.ok(declaredIntent, `${label}: declared intent`);
    assert.equal(intent.targetPage, declaredIntent.targetPage, `${label}: target page`);
    assert.equal(intent.requiredRuns, 3, `${label}: required runs`);
    assert.equal(intent.completedRuns, 3, `${label}: completed runs`);
    assert.deepEqual(intent.runs.map(({ repetition }) => repetition), [1, 2, 3], `${label}: clean repetitions`);
    const passing = intent.runs.filter(({ status }) => status === "passed");
    assert.equal(intent.passedRuns, passing.length, `${label}: passed-run counter`);
    passedRuns += passing.length;
    if (intent.deterministic) deterministicIntents += 1;
    for (const run of intent.runs) {
      const runLabel = `${label}/${run.repetition}`;
      assert.equal(run.sourceSha256After, source.sourceSha256, `${runLabel}: source hash preserved`);
      assert.equal(run.worker.schema, "office-kit/pptx-programmable-import-worker/v1", `${runLabel}: public worker receipt`);
      assert.equal(run.worker.sourceUnchanged, true, `${runLabel}: source unmodified`);
      assert.equal(run.worker.secondImport, true, `${runLabel}: second import`);
      assert.equal(run.worker.outputSha256, run.outputSha256, `${runLabel}: output receipt hash`);
      assert.equal(run.worker.observedValue, declaredIntent.value, `${runLabel}: intended value re-imported`);
      assert.ok(new Set(["presentation.editNativeLeaf", "ImageElement.editSvgText"]).has(run.worker.publicApi), `${runLabel}: public API only`);
      assert.equal(run.packageOracle.sourceSha256, source.sourceSha256, `${runLabel}: package source identity`);
      assert.equal(run.packageOracle.outputSha256, run.outputSha256, `${runLabel}: package output identity`);
      assert.equal(run.packageOracle.partSet.passed, true, `${runLabel}: OPC part set`);
      assert.equal(run.packageOracle.nonTargetPartsByteIdentical, true, `${runLabel}: non-target OPC bytes`);
      assert.equal(run.packageOracle.relationships.passed, true, `${runLabel}: relationships`);
      assert.equal(run.packageOracle.targetMask.passed, true, `${runLabel}: masked XML/SVG`);
      if (run.packageOracle.nestedPackage) assert.equal(run.packageOracle.nestedPackage.passed, true, `${runLabel}: nested package`);
      if (run.status === "failed") {
        assert.match(run.reason, /^Target rendered page \d+ did not change$/u, `${runLabel}: real pixel failure retained`);
        assert.equal(run.pixelOracle.passed, false, `${runLabel}: failed pixel result retained`);
        assert.equal(run.pixelOracle.reason, run.reason, `${runLabel}: pixel reason retained exactly`);
        continue;
      }
      assert.equal(run.status, "passed", `${runLabel}: valid status`);
      assert.equal(run.pixelOracle.passed, true, `${runLabel}: pixel oracle`);
      assert.equal(run.pixelOracle.targetPageChanged, true, `${runLabel}: target page changed`);
      assert.equal(run.pixelOracle.nonTargetPagesPixelIdentical, true, `${runLabel}: non-target pixels`);
      assert.deepEqual(run.pixelOracle.nonTargetMismatches, [], `${runLabel}: no non-target pixel drift`);
    }
    const outputHashes = new Set(passing.map(({ outputSha256 }) => outputSha256));
    if (intent.deterministic) assert.equal(outputHashes.size, 1, `${label}: byte deterministic`);
    if (!intent.deterministic && passing.length === 3) assert.equal(outputHashes.size, 3, `${label}: nondeterminism retained`);
  }
}

assert.equal(passedRuns, 60);
assert.equal(deterministicIntents, 10);
assert.deepEqual(matrix.acceptance, {
  requiredIntents: 30,
  requiredRuns: 90,
  passedRuns: 60,
  deterministicIntents: 10,
  status: "failed",
});
assert.deepEqual(matrix.replay, {
  mode: "independent-evaluator-only",
  originalEvidenceSha256: "42d364d94997a249e94750adb0b72faebbe62f08cc99ba7bcb6585bb3ebc48ac",
  editsRerun: false,
  outcomesChanged: false,
  partialChecksRetained: true,
});

const rebuilt = buildBaselineEvidence({ matrix, codex, matrixBytes, codexBytes, harnessHead: baseline.acceptanceHarnessHead });
assert.deepEqual(baseline, rebuilt, "summary must be derived exactly from immutable component evidence");
assert.equal(baseline.productBaseline, BASELINE);
assert.equal(baseline.evidenceFiles.matrix.sha256, sha256(matrixBytes));
assert.equal(baseline.evidenceFiles.codex.sha256, sha256(codexBytes));
assert.equal(baseline.matrix.nonDeterministic.length, 20);
assert.equal(baseline.matrix.nonDeterministic.filter(({ passedRuns: count }) => count === 0).length, 10);
assert.equal(baseline.matrix.nonDeterministic.filter(({ passedRuns: count }) => count === 3).length, 10);
assert.equal(baseline.codex.failedTrials.length, 9);
assert.ok(baseline.codex.failedTrials.every(({ failures, agentFinal, policyPassed, outputPassed }) => failures.length > 0 && agentFinal && policyPassed && !outputPassed));
assert.deepEqual(baseline.acceptance, {
  status: "failed",
  failuresPreserved: true,
  oracleWeakened: false,
  productModifiedByAcceptance: false,
});
assert.doesNotMatch(Buffer.concat([matrixBytes, codexBytes, baselineBytes]).toString("utf8"), /(?:\/Users\/|\/tmp\/)/u, "evidence must be machine portable");

console.log("PPTX programmable-import baseline evidence ok (60/90 runs, 10/30 deterministic, 0/9 Codex)");
