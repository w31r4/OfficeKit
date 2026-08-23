import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const runParent = await mkdtemp(path.join(os.tmpdir(), "officekit-pptx-source-derived-companion-test-"));
const runRoot = path.join(runParent, "run");

try {
  const result = spawnSync(process.execPath, [
    "scripts/pptx-source-derived-companion.mjs",
    "--run-root", runRoot,
    "--case", "controlled-clone-table",
    "--repetitions", "2",
    "--no-render",
  ], { cwd: repoRoot, encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
  assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);

  const evidence = JSON.parse(await readFile(path.join(runRoot, "evidence.json"), "utf8"));
  assert.equal(evidence.schema, "office-kit/pptx-source-derived-companion-evidence/v1");
  assert.deepEqual(evidence.acceptance, { scope: "selected-case", status: "passed" });
  assert.equal(evidence.coverage.status, "partial");
  assert.deepEqual(evidence.coverage.passed, ["add", "table"]);
  const item = assertSingle(evidence.cases);
  assert.equal(item.id, "controlled-clone-table");
  assert.equal(item.requiredRuns, 2);
  assert.equal(item.passedRuns, 2);
  assert.equal(item.deterministic, true);
  assert.equal(new Set(item.runs.map(({ outputSha256 }) => outputSha256)).size, 1);
  for (const run of item.runs) {
    assert.equal(run.worker.sourceUnchanged, true);
    assert.equal(run.worker.secondImport.passed, true);
    assert.equal(run.packageOracle.passed, true);
    assert.equal(run.packageOracle.nonTargetPartsByteIdentical, true);
    assert.deepEqual(run.packageOracle.partSet, { passed: true, added: [], removed: [] });
    assert.equal(run.packageOracle.changedParts.length, 1);
    assert.match(run.packageOracle.changedParts[0], /^ppt\/slides\/slide\d+[.]xml$/u);
    assert.equal(run.packageOracle.targetMask.passed, true);
    assert.equal(run.pixelOracle.skipped, true);
  }
  assert.equal(assertSingle(evidence.existingEvidence).passed, true);

  const collision = spawnSync(process.execPath, [
    "scripts/pptx-source-derived-companion.mjs",
    "--run-root", runRoot,
    "--case", "controlled-clone-table",
    "--repetitions", "1",
    "--no-render",
  ], { cwd: repoRoot, encoding: "utf8" });
  assert.notEqual(collision.status, 0);
  assert.match(collision.stderr, /outputs are create-only/u);
} finally {
  await rm(runParent, { recursive: true, force: true });
}

console.log("PPTX source-derived companion smoke ok");

function assertSingle(values) {
  assert.equal(values.length, 1);
  return values[0];
}
