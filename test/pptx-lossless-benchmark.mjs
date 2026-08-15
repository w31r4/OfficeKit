import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";

const manifestPath = path.resolve("evals/pptx-lossless/manifest.v1.json");
const manifestBytes = await readFile(manifestPath);
const manifest = JSON.parse(manifestBytes.toString("utf8"));
const evidence = JSON.parse(await readFile(path.resolve("evals/pptx-lossless/evidence.v1.json"), "utf8"));

assert.equal(manifest.schema, "office-kit/pptx-lossless-benchmark/v1");
assert.equal(manifest.sources.length, 3);
assert.equal(new Set(manifest.sources.map((source) => source.id)).size, 3);
assert.equal(new Set(manifest.sources.map((source) => source.sha256)).size, 3);
assert.equal(manifest.sources.reduce((count, source) => count + source.targets.length, 0), 10);

for (const source of manifest.sources) {
  assert.match(source.id, /^[a-z0-9][a-z0-9-]+$/u);
  assert.equal(path.isAbsolute(source.fileName), false);
  assert.equal(source.fileName.includes(".."), false);
  assert.match(source.sha256, /^[a-f0-9]{64}$/u);
  assert.equal(Number.isSafeInteger(source.bytes) && source.bytes > 0, true);
  assert.equal(source.inventory.partCount, Object.keys(source.inventory.partHashes).length);
  assert.equal(Array.isArray(source.editableNodes), true);
  assert.equal("sourceBytes" in source, false);
  for (const [partPath, sha256] of Object.entries(source.inventory.partHashes)) {
    assert.equal(path.posix.isAbsolute(partPath), false);
    assert.equal(partPath.split("/").includes(".."), false);
    assert.match(sha256, /^[a-f0-9]{64}$/u);
  }
  for (const target of source.targets) {
    assert.match(target.nodeId, /^presentation\/slide\/[1-9][0-9]*\/element\/[1-9][0-9]*(?:\/.*)?$/u);
    if ((target.operation ?? "text") === "text") {
      assert.equal(typeof target.expected, "string");
      assert.equal(typeof target.value, "string");
      assert.equal(target.search === undefined || typeof target.search === "string", true);
      assert.equal(target.result === undefined || typeof target.result === "string", true);
      assert.notEqual(target.search ?? target.expected, target.value);
      assert.equal(source.editableNodes.some((node) => node.id === target.nodeId && node.text === target.expected), true);
    } else {
      assert.match(target.operation, /^native(?:Leaf|Leaves)$/u);
      const leaves = target.operation === "nativeLeaves" ? target.leaves : [target];
      assert.equal(Array.isArray(leaves) && leaves.length > 0 && leaves.length <= 8, true);
      assert.equal(new Set(leaves.map((leaf) => leaf.leafKind)).size, leaves.length);
      for (const leaf of leaves) {
        assert.match(leaf.leafKind, /^(text|chartTitleText|leftEmu|topEmu|widthEmu|heightEmu)$/u);
        if (leaf.leafKind === "text" || leaf.leafKind === "chartTitleText") {
          assert.equal(typeof leaf.expectedValue, "string");
          assert.equal(typeof leaf.value, "string");
        } else {
          assert.equal(Number.isSafeInteger(leaf.expectedValue), true);
          assert.equal(Number.isSafeInteger(leaf.value), true);
        }
        assert.notEqual(leaf.expectedValue, leaf.value);
      }
    }
  }
}

assert.equal(
  manifest.sources.some((source) => source.targets.some((target) =>
    source.editableNodes.some((node) => node.id === target.nodeId && node.depth > 0))),
  true,
);

assert.equal(evidence.schema, "office-kit/pptx-lossless-evidence/v1");
assert.equal(evidence.manifestSha256, createHash("sha256").update(manifestBytes).digest("hex"));
assert.equal(evidence.repetitionsPerTarget, 3);
assert.deepEqual(Object.values(evidence.runnerContract), [true, true, true, true, true, true]);
assert.equal(evidence.sources.length, manifest.sources.length);
assert.equal(evidence.sources.reduce((count, source) => count + source.targets.length, 0), 10);
for (const source of evidence.sources) {
  const declared = manifest.sources.find((candidate) => candidate.id === source.id);
  assert.ok(declared);
  assert.equal(source.sourceSha256, declared.sha256);
  assert.equal(source.noOpByteIdentical, true);
  assert.equal(source.targets.length, declared.targets.length);
  for (const target of source.targets) {
    assert.equal(declared.targets.some((candidate) => candidate.id === target.id), true);
    assert.match(target.outputSha256, /^[a-f0-9]{64}$/u);
    assert.equal(target.changedParts.length, 1);
    assert.match(target.changedParts[0], /^ppt\/(?:slides\/slide[1-9][0-9]*|charts\/chart[1-9][0-9]*)\.xml$/u);
  }
}

console.log("PPTX lossless benchmark manifest smoke ok");
