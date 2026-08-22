import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const evidencePath = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../evals/pptx-generation/agent-blackbox.v1.json");
const evidence = JSON.parse(await readFile(evidencePath, "utf8"));

assert.equal(evidence.schema, "office-kit/pptx-template-conditioned-generation-agent-blackbox/v1");
assert.equal(evidence.package.version, "0.6.0");
assert.match(evidence.package.sha256, /^[a-f0-9]{64}$/u);
assert.equal(evidence.acceptance.required, 3);
assert.equal(evidence.acceptance.completed, 3);
assert.equal(evidence.acceptance.allReimported, true);
assert.equal(evidence.acceptance.allSourceProtected, true);
assert.equal(evidence.acceptance.allPackageNonTargetPreserved, true);
assert.equal(evidence.acceptance.allAuditsPresent, true);
assert.equal(evidence.acceptance.status, "passed");

const expected = new Map([
  ["suanzhi-future-2026", 21],
  ["blue-gray-acid-template", 19],
  ["mckinsey-customer-loyalty", 8],
]);
assert.equal(evidence.trials.length, expected.size);
for (const trial of evidence.trials) {
  assert.ok(expected.has(trial.sourceId), `unexpected black-box source ${trial.sourceId}`);
  assert.equal(trial.sourceSlides, expected.get(trial.sourceId));
  assert.match(trial.sourceSha256, /^[a-f0-9]{64}$/u);
  assert.match(trial.outputSha256, /^[a-f0-9]{64}$/u);
  assert.ok(trial.outputSlides > trial.sourceSlides);
  assert.ok(trial.generatedSlides >= 1);
  assert.equal(trial.sourceProtected, true);
  assert.equal(trial.sourceSlidesPreserved, true);
  assert.equal(trial.packageOracle.nonTargetPartsByteIdentical, true);
  assert.deepEqual(trial.packageOracle.changed, []);
  assert.deepEqual(trial.packageOracle.missing, []);
  assert.equal(trial.reimport.passed, true);
  assert.equal(trial.profileCaptured, true);
  assert.equal(trial.agent.exitStatus, 0);
  assert.deepEqual(trial.agent.phases.map(({ phase, status }) => ({ phase, status })), [
    { phase: "plan", status: 0 },
    { phase: "author", status: 0 },
    { phase: "review", status: 0 },
  ]);
  assert.equal(trial.audit.hasFrameMap, true);
  assert.equal(trial.audit.hasProfileSummary, true);
  assert.equal(trial.audit.sourceProtected, true);
  assert.equal(trial.audit.reimportOk, true);
  assert.ok(["unavailable", "complete", "requires-human"].includes(trial.audit.visualReview));
}

console.log("pptx template-conditioned Agent black-box evidence smoke ok");
