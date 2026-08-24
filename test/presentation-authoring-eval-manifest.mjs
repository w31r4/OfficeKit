import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const manifestPath = path.join(repoRoot, "evals", "presentation-authoring-compiler", "pilot.v1.json");
const manifest = JSON.parse(await fs.readFile(manifestPath, "utf8"));

assert.equal(manifest.schema, "office-kit/presentation-authoring-pilot/v1");
assert.equal(manifest.status, "not-run");
assert.equal(manifest.baseline.arm, "A");
assert.equal(manifest.baseline.route, "grid-default");
assert.equal(manifest.baseline.frozenAt, "origin/main@a0452867");
assert.equal(manifest.baseline.defaultUntilThresholdsPass, true);
assert.deepEqual(Object.keys(manifest.arms).sort(), ["A", "B", "C"]);
assert.equal(manifest.tasks.length, 10);
assert.equal(new Set(manifest.tasks.map((task) => task.scenario)).size, 5);
assert.equal(manifest.design.trialsPerTask * manifest.design.armsPerTask * manifest.tasks.length, manifest.design.totalRuns);
assert.equal(manifest.design.sameInputsAcrossArms, true);
assert.equal(manifest.design.freshContextPerRun, true);
assert.equal(manifest.design.package, "packed clean-install");
for (const threshold of Object.values(manifest.thresholds)) assert.ok([">=", "<=", ">"].includes(threshold.operator));
assert.match(manifest.rolloutRule, /every threshold passes.*otherwise keep A shipped.*C.*experimental/i);
assert.doesNotMatch(JSON.stringify(manifest), /(?:\/Users\/|[A-Z]:\\|\/tmp\/)/u);

console.log("presentation authoring evaluation manifest ok");
