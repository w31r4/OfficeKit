import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const packageJson = JSON.parse(await fs.readFile(path.join(repoRoot, "package.json"), "utf8"));
const gateRunner = await fs.readFile(path.join(repoRoot, "scripts", "run-test-gate.mjs"), "utf8");
const ci = await fs.readFile(path.join(repoRoot, ".github", "workflows", "ci.yml"), "utf8");
const slow = await fs.readFile(path.join(repoRoot, ".github", "workflows", "ci-slow.yml"), "utf8");
const windows = await fs.readFile(path.join(repoRoot, ".github", "workflows", "windows-office-live.yml"), "utf8");

assert.equal(packageJson.scripts.test, "node scripts/run-test-gate.mjs fast");
assert.equal(packageJson.scripts["test:fast"], packageJson.scripts.test);
assert.equal(packageJson.scripts["test:slow"], "node scripts/run-test-gate.mjs slow");
assert.equal(packageJson.scripts["test:slow:templates"], "node test/default-template-library.mjs");
assert.equal(packageJson.scripts["test:slow:promptbench"], "node test/agent-evals.mjs");

const fastStart = gateRunner.indexOf("const fastSteps");
const slowStart = gateRunner.indexOf("const slowSteps");
const fastSource = gateRunner.slice(fastStart, slowStart);
const slowSource = gateRunner.slice(slowStart, gateRunner.indexOf("const steps"));
assert.match(fastSource, /check-js-syntax/);
assert.match(fastSource, /windows-live-gate/);
assert.doesNotMatch(fastSource, /default-template-library|agent-evals|pdf-provider-pack-build|pdf-pyhanko-provider|presentation-skill|document-skill|pdf-skill|office-kit-package/);
assert.match(slowSource, /default-template-library/);
assert.match(slowSource, /agent-evals/);
assert.match(slowSource, /pdf-provider-pack-build/);

assert.match(ci, /name:\s*Fast gate/);
assert.match(ci, /npm test/);
assert.doesNotMatch(ci, /Install isolated PyMuPDF|default-template-library|agent-evals/);
assert.match(slow, /workflow_dispatch:/);
assert.match(slow, /schedule:/);
assert.match(slow, /npm run test:slow/);
assert.match(slow, /timeout-minutes:\s*45/);
assert.match(windows, /workflow_dispatch:/);
assert.match(windows, /self-hosted/);
assert.match(windows, /validate-windows-live-evidence/);
assert.doesNotMatch(windows, /test:excel-live|test:powerpoint-live/);

console.log("gate policy contract ok");
