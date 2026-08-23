import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

const root = path.resolve(import.meta.dirname, "..");
const intents = JSON.parse(await readFile(path.join(root, "evals/pptx-programmable-import/intent-matrix.v1.json"), "utf8"));
const continuations = JSON.parse(await readFile(path.join(root, "evals/pptx-programmable-import/continuation-tasks.v1.json"), "utf8"));

assert.equal(intents.schema, "office-kit/pptx-programmable-import-intents/v1");
assert.equal(intents.baseline, "d5df8df94727dccd4412e6be874d1c5407b57f64");
assert.equal(intents.repetitionsPerIntent, 3);
assert.equal(intents.cleanSourcePerRun, true);
assert.equal(intents.sources.length, 3);
assert.deepEqual(intents.sources.map(({ id }) => id), [
  "suanzhi-future-2026",
  "blue-gray-acid-template",
  "mckinsey-customer-loyalty",
]);

const sourceIds = new Set();
const intentIds = new Set();
for (const source of intents.sources) {
  assert.match(source.sha256, /^[a-f0-9]{64}$/u);
  assert.ok(Number.isInteger(source.slideCount) && source.slideCount > 0);
  assert.ok(Array.isArray(source.intents) && source.intents.length >= 10, `${source.id} needs at least ten intents`);
  assert.equal(sourceIds.has(source.id), false);
  sourceIds.add(source.id);
  for (const intent of source.intents) {
    const key = `${source.id}/${intent.id}`;
    assert.equal(intentIds.has(key), false, `duplicate intent ${key}`);
    intentIds.add(key);
    assert.ok(["native-leaf", "svg-text"].includes(intent.operation));
    assert.match(intent.targetId, /^presentation\/slide\/\d+\/element\/\d+/u);
    assert.ok(Number.isInteger(intent.targetPage) && intent.targetPage >= 1 && intent.targetPage <= source.slideCount);
    assert.notDeepEqual(intent.expected, intent.value);
    assert.ok(Array.isArray(intent.oracle?.changedParts) && intent.oracle.changedParts.length > 0);
    assert.equal(new Set(intent.oracle.changedParts).size, intent.oracle.changedParts.length);
    for (const part of [...intent.oracle.changedParts, ...(intent.oracle.addedParts || [])]) {
      assert.equal(path.posix.isAbsolute(part), false);
      assert.equal(part.split("/").includes(".."), false);
    }
    if (intent.operation === "svg-text") {
      assert.match(intent.nodeId, /^svg-text-\d+$/u);
      assert.match(intent.oracle.sourceSvgPart, /^ppt\/media\/image\d+[.]svg$/u);
      assert.deepEqual(intent.oracle.addedParts, ["ppt/media/image9.svg"]);
    } else {
      assert.equal(typeof intent.leafKind, "string");
      assert.equal(intent.oracle.addedParts, undefined);
    }
  }
}
assert.equal(intentIds.size, 30);

assert.equal(continuations.schema, "office-kit/pptx-codex-continuation-tasks/v1");
assert.equal(continuations.baseline, intents.baseline);
assert.equal(continuations.trialsPerTask, 3);
assert.equal(continuations.freshCodexContextPerTrial, true);
assert.equal(continuations.replSessionsPerTrial, 3);
assert.equal(continuations.tasks.length, 3);
assert.deepEqual(continuations.tasks.map(({ sourceId }) => sourceId), intents.sources.map(({ id }) => id));
assert.equal(new Set(continuations.tasks.map(({ id }) => id)).size, 3);
assert.deepEqual(continuations.tasks.map(({ acceptanceRenderer }) => acceptanceRenderer), [
  "keynote",
  "keynote",
  "libreoffice",
]);
for (const task of continuations.tasks) {
  const source = intents.sources.find(({ id }) => id === task.sourceId);
  assert.ok(source);
  assert.match(task.output, /^outputs\/[a-z0-9-]+[.]pptx$/u);
  assert.ok(Number.isInteger(task.sourceSlide) && task.sourceSlide >= 1 && task.sourceSlide <= source.slideCount);
  assert.equal(task.targetPageAfterAppend, source.slideCount + 1);
  assert.ok(task.goal.length >= 100);
  assert.equal(task.expectedTexts.length, 2);
  assert.ok(task.expectedTexts.every((value) => typeof value === "string" && value.length > 0));
  assert.equal(task.edits.length, 2);
  assert.deepEqual(task.edits.map(({ phase }) => phase), [1, 2]);
  assert.deepEqual(task.edits.map(({ value }) => value), task.expectedTexts);
  for (const edit of task.edits) {
    assert.ok(["native-leaf", "svg-text"].includes(edit.operation));
    assert.match(edit.targetId, new RegExp(`^presentation/slide/${task.targetPageAfterAppend}/element/\\d+$`, "u"));
    assert.equal(typeof edit.expected, "string");
    assert.equal(typeof edit.value, "string");
    assert.notEqual(edit.expected, edit.value);
    if (edit.operation === "native-leaf") {
      assert.equal(edit.leafKind, "text");
      assert.ok(Number.isInteger(edit.textLeafIndex));
    } else {
      assert.match(edit.nodeId, /^svg-text-\d+$/u);
    }
  }
}

console.log("PPTX programmable-import task definitions ok");
