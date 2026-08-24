import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

import { buildCandidateEvidence } from "../scripts/pptx-programmable-import-baseline.mjs";
import { sha256 } from "../scripts/pptx-programmable-import-oracle.mjs";

const PRODUCT_BASELINE = "d5df8df94727dccd4412e6be874d1c5407b57f64";
const CANDIDATE_HEAD = "e1bb8699671c3599b44b999ca308ff8d0d9581d7";
const TARBALL_SHA256 = "0152742d17a07a7b53e53f83f75c08c829804ab8f73ad65841a5e49946e7e8a9";
const root = path.resolve(import.meta.dirname, "../evals/pptx-programmable-import");
const paths = {
  matrix: path.join(root, "candidate/matrix.v1.json"),
  codex: path.join(root, "candidate/codex.v1.json"),
  companion: path.join(root, "source-derived-companion.evidence.v1.json"),
  candidate: path.join(root, "candidate.v1.json"),
  baseline: path.join(root, "baseline.v1.json"),
  intents: path.join(root, "intent-matrix.v1.json"),
  continuation: path.join(root, "continuation-tasks.v1.json"),
  companionDefinitions: path.join(root, "source-derived-companion.v1.json"),
};
const bytes = Object.fromEntries(await Promise.all(Object.entries(paths).map(async ([name, file]) => [name, await readFile(file)])));
const parsed = Object.fromEntries(Object.entries(bytes).map(([name, value]) => [name, JSON.parse(value)]));

assert.equal(parsed.baseline.acceptance.status, "failed", "historical failed baseline must remain preserved");
assert.equal(parsed.matrix.baseline, PRODUCT_BASELINE);
assert.equal(parsed.codex.baseline, PRODUCT_BASELINE);
assert.equal(parsed.companion.productBaseline, PRODUCT_BASELINE);
assert.equal(parsed.candidate.productBaseline, PRODUCT_BASELINE);
assert.equal(parsed.candidate.candidateHead, CANDIDATE_HEAD);
assert.equal(parsed.matrix.definitionsSha256, sha256(bytes.intents));
assert.equal(parsed.codex.definitions.intentSha256, sha256(bytes.intents));
assert.equal(parsed.codex.definitions.continuationSha256, sha256(bytes.continuation));
assert.equal(parsed.companion.definitionsSha256, sha256(bytes.companionDefinitions));

assert.equal(parsed.matrix.package.tarballSha256, TARBALL_SHA256);
assert.equal(parsed.codex.package.tarballSha256, TARBALL_SHA256);
assert.equal(parsed.matrix.package.installKind, "packed-clean-install");
assert.equal(parsed.codex.package.cleanInstallPerTrial, true);
assert.deepEqual(parsed.matrix.acceptance, {
  requiredIntents: 30,
  requiredRuns: 90,
  passedRuns: 90,
  deterministicIntents: 30,
  status: "passed",
});
assert.deepEqual(parsed.codex.acceptance, {
  requiredTasks: 3,
  trialsPerTask: 3,
  requiredTrials: 9,
  completedTrials: 9,
  passedTrials: 9,
  status: "passed",
});

const definitionsBySource = new Map(parsed.intents.sources.map((source) => [source.id, source]));
const rendererCounts = new Map();
for (const source of parsed.matrix.sources) {
  const definition = definitionsBySource.get(source.id);
  assert.ok(definition, `${source.id}: declared source`);
  assert.equal(source.sourceSha256, definition.sha256, `${source.id}: immutable source hash`);
  assert.equal(source.intents.length, 10, `${source.id}: ten intents`);
  for (const intent of source.intents) {
    const declared = definition.intents.find(({ id }) => id === intent.id);
    const label = `${source.id}/${intent.id}`;
    assert.ok(declared, `${label}: declared intent`);
    assert.equal(intent.targetPage, declared.targetPage, `${label}: target page`);
    assert.equal(intent.passedRuns, 3, `${label}: three passing runs`);
    assert.equal(intent.deterministic, true, `${label}: deterministic`);
    assert.deepEqual(intent.runs.map(({ repetition }) => repetition), [1, 2, 3], `${label}: independent repetitions`);
    assert.equal(new Set(intent.runs.map(({ outputSha256 }) => outputSha256)).size, 1, `${label}: byte-identical output`);
    for (const run of intent.runs) {
      const runLabel = `${label}/${run.repetition}`;
      assert.equal(run.status, "passed", `${runLabel}: passed`);
      assert.equal(run.sourceSha256After, source.sourceSha256, `${runLabel}: source unchanged`);
      assert.equal(run.worker.observedValue, declared.value, `${runLabel}: intended value reimported`);
      assert.equal(run.worker.secondImport, true, `${runLabel}: second import`);
      assert.equal(run.packageOracle.partSet.passed, true, `${runLabel}: part set`);
      assert.equal(run.packageOracle.nonTargetPartsByteIdentical, true, `${runLabel}: non-target bytes`);
      assert.equal(run.packageOracle.relationships.passed, true, `${runLabel}: relationships`);
      assert.equal(run.packageOracle.targetMask.passed, true, `${runLabel}: masked target`);
      if (run.packageOracle.nestedPackage) assert.equal(run.packageOracle.nestedPackage.passed, true, `${runLabel}: nested package`);
      assert.equal(run.pixelOracle.passed, true, `${runLabel}: pixel oracle`);
      assert.equal(run.pixelOracle.targetPageChanged, true, `${runLabel}: target page changed`);
      assert.equal(run.pixelOracle.nonTargetPagesPixelIdentical, true, `${runLabel}: non-target pages`);
      assert.deepEqual(run.pixelOracle.nonTargetMismatches, [], `${runLabel}: no non-target pixel drift`);
      rendererCounts.set(run.pixelOracle.renderer, (rendererCounts.get(run.pixelOracle.renderer) || 0) + 1);
    }
  }
}
assert.deepEqual(Object.fromEntries([...rendererCounts].sort()), { keynote: 30, libreoffice: 60 });

const trialsByTask = Map.groupBy(parsed.codex.trials, ({ taskId }) => taskId);
assert.equal(trialsByTask.size, 3);
for (const [taskId, trials] of trialsByTask) {
  assert.deepEqual(trials.map(({ repetition }) => repetition).sort(), [1, 2, 3], `${taskId}: three fresh trials`);
  assert.equal(new Set(trials.map(({ checks }) => checks.durableTask.publication.sha256)).size, 1, `${taskId}: byte-deterministic publication`);
  for (const trial of trials) {
    const label = `${taskId}/${trial.repetition}`;
    assert.equal(trial.freshCodexContext, true, `${label}: fresh Codex context`);
    assert.equal(trial.packedCleanInstall.tarballSha256, TARBALL_SHA256, `${label}: exact package`);
    assert.equal(trial.checks.policy.findings.length, 0, `${label}: allowed paths only`);
    assert.equal(trial.checks.source.passed, true, `${label}: source protected`);
    assert.equal(trial.checks.output.createOnly, true, `${label}: create-only output`);
    assert.equal(trial.checks.durableTask.sessions, 3, `${label}: three sessions`);
    assert.deepEqual(trial.checks.durableTask.commits.map(({ commitId }) => commitId), ["c0001", "c0002"], `${label}: two reviewed commits`);
    assert.equal(trial.checks.durableTask.publication.commitId, "c0002", `${label}: sole final publication`);
    assert.equal(trial.checks.packageOracle.partSet.passed, true, `${label}: package part set`);
    assert.equal(trial.checks.packageOracle.relationships.passed, true, `${label}: package relationships`);
    assert.equal(trial.checks.packageOracle.targetMask.passed, true, `${label}: package target mask`);
    assert.equal(trial.checks.secondImport.passed, true, `${label}: second import`);
    assert.equal(trial.checks.pixelOracle.passed, true, `${label}: visual oracle`);
    assert.equal(trial.checks.pixelOracle.nonTargetPagesPixelIdentical, true, `${label}: non-target pixels`);
  }
}

const rebuilt = buildCandidateEvidence({
  matrix: parsed.matrix,
  codex: parsed.codex,
  companion: parsed.companion,
  matrixBytes: bytes.matrix,
  codexBytes: bytes.codex,
  companionBytes: bytes.companion,
  candidateHead: CANDIDATE_HEAD,
});
assert.deepEqual(parsed.candidate, rebuilt, "candidate summary must derive exactly from component evidence");
assert.deepEqual(parsed.candidate.acceptance, {
  status: "passed",
  failuresPreserved: true,
  oracleWeakened: false,
  productModifiedByAcceptance: false,
});
assert.equal(parsed.candidate.evidenceFiles.matrix.sha256, sha256(bytes.matrix));
assert.equal(parsed.candidate.evidenceFiles.codex.sha256, sha256(bytes.codex));
assert.equal(parsed.candidate.evidenceFiles.companion.sha256, sha256(bytes.companion));
assert.doesNotMatch(Buffer.concat([bytes.matrix, bytes.codex, bytes.companion, bytes.candidate]).toString("utf8"), /(?:\/Users\/|\/tmp\/)/u, "candidate evidence must be machine portable");

console.log("PPTX programmable-import candidate evidence ok (90/90 matrix, 9/9 Codex, 24/24 continuation)");
