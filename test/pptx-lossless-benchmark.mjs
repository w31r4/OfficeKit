import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

const manifestPath = path.resolve("evals/pptx-lossless/manifest.v1.json");
const manifest = JSON.parse(await readFile(manifestPath, "utf8"));

assert.equal(manifest.schema, "office-kit/pptx-lossless-benchmark/v1");
assert.equal(manifest.sources.length, 3);
assert.equal(new Set(manifest.sources.map((source) => source.id)).size, 3);
assert.equal(new Set(manifest.sources.map((source) => source.sha256)).size, 3);
assert.equal(manifest.sources.reduce((count, source) => count + source.targets.length, 0), 6);

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
    assert.equal(typeof target.expected, "string");
    assert.equal(typeof target.value, "string");
    assert.notEqual(target.expected, target.value);
    assert.equal(source.editableNodes.some((node) => node.id === target.nodeId && node.text === target.expected), true);
  }
}

console.log("PPTX lossless benchmark manifest smoke ok");
