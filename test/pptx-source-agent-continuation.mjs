import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

const evidence = JSON.parse(await readFile(path.resolve("evals/pptx-lossless/source-agent-continuation.v1.json"), "utf8"));
assert.equal(evidence.schema, "office-kit/pptx-source-agent-continuation-rehearsal/v1");
assert.deepEqual(evidence.protocol, { repl: 2, visualReview: "unavailable", package: "public-office-kit" });
assert.equal(evidence.modelBlackBox.required, 3);
assert.equal(evidence.modelBlackBox.completed, 0);
assert.equal(evidence.modelBlackBox.status, "open");
assert.deepEqual(Object.keys(evidence.modelBlackBox).sort(), ["completed", "required", "status"]);
assert.deepEqual(evidence.sources.map((source) => source.id), [
  "suanzhi-future-2026",
  "blue-gray-acid-template",
  "mckinsey-customer-loyalty",
]);
for (const source of evidence.sources) {
  assert.equal(source.taskId, "t_<redacted>");
  assert.equal(source.taskIdValidated, true);
  assert.equal(source.freshSessions, 3);
  assert.equal(source.sourceUnchanged, true);
  assert.equal(source.commits.length, 2);
  assert.deepEqual(source.commits.map((commit) => commit.commitId), ["c0001", "c0002"]);
  assert.ok(source.commits.every((commit) => commit.reviewVerdict === "passed-with-limitations" && commit.visualReview === "unavailable"));
  assert.equal(source.finalVerification.slideCount, source.sourceSlideCount + 1);
  assert.equal(source.finalVerification.result.foundResumed, true);
  assert.equal(source.publishedPathRelative, `outputs/${source.id}-continued.pptx`);
  assert.match(source.sourceSha256, /^[a-f0-9]{64}$/u);
  assert.match(source.publishedSha256, /^[a-f0-9]{64}$/u);
  assert.equal(source.trace.length, 5);
  assert.ok(source.trace.every((entry) => entry.ok === true && entry.maybeApplied === false));
}
console.log("PPTX source-derived continuation REPL rehearsal evidence ok");
