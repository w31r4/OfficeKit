import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const evidencePath = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../evals/pptx-generation/agent-multipage.v2.json");
const evidence = JSON.parse(await readFile(evidencePath, "utf8"));

assert.equal(evidence.schema, "office-kit/pptx-template-conditioned-generation-agent-multipage/v2");
assert.equal(evidence.package.version, "0.6.0");
assert.match(evidence.package.sha256, /^[a-f0-9]{64}$/u);
assert.equal(evidence.mode, "multi");
assert.equal(evidence.requestedPages, 10);
assert.deepEqual(evidence.acceptance, {
  required: 3,
  completed: 3,
  allPlanned: true,
  allGenerated: true,
  allReimported: true,
  allSourceProtected: true,
  allPackageNonTargetPreserved: true,
  allAuditsPresent: true,
  status: "passed",
});

const expected = new Map([
  ["suanzhi-future-2026", 21],
  ["blue-gray-acid-template", 19],
  ["mckinsey-customer-loyalty", 8],
]);
assert.equal(evidence.trials.length, expected.size);
for (const trial of evidence.trials) {
  assert.ok(expected.has(trial.sourceId), `unexpected multipage source ${trial.sourceId}`);
  assert.equal(trial.sourceSlides, expected.get(trial.sourceId));
  assert.match(trial.sourceSha256, /^[a-f0-9]{64}$/u);
  assert.match(trial.outputSha256, /^[a-f0-9]{64}$/u);
  assert.equal(trial.generatedSlides, 10);
  assert.equal(trial.outputSlides, trial.sourceSlides + 10);
  assert.equal(trial.plannedPages, 10);
  assert.equal(trial.sourceProtected, true);
  assert.equal(trial.sourceSlidesPreserved, true);
  assert.equal(trial.packageOracle.nonTargetPartsByteIdentical, true);
  assert.deepEqual(trial.packageOracle.changed, []);
  assert.deepEqual(trial.packageOracle.missing, []);
  assert.equal(trial.reimport.passed, true);
  assert.equal(trial.agent.exitStatus, 0);
  assert.deepEqual(trial.agent.phases.map(({ phase, status }) => ({ phase, status })), [
    { phase: "plan", status: 0 },
    { phase: "author", status: 0 },
    { phase: "review", status: 0 },
  ]);
  assert.equal(trial.audit.hasFrameMap, true);
  assert.equal(trial.audit.hasProfileSummary, true);
  assert.equal(trial.audit.sourceProtected, true);
  assert.equal(trial.audit.reimportPassed, true);
  assert.equal(trial.audit.visualReview, "unavailable");
}

console.log("pptx multipage Agent evidence smoke ok");
