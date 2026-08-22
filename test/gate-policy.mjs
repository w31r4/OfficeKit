import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const packageJson = JSON.parse(await fs.readFile(path.join(repoRoot, "package.json"), "utf8"));
const gateRunner = await fs.readFile(path.join(repoRoot, "scripts", "run-test-gate.mjs"), "utf8");
const ci = await fs.readFile(path.join(repoRoot, ".github", "workflows", "ci.yml"), "utf8");
const slow = await fs.readFile(path.join(repoRoot, ".github", "workflows", "ci-slow.yml"), "utf8");
const release = await fs.readFile(path.join(repoRoot, ".github", "workflows", "release.yml"), "utf8");
const windows = await fs.readFile(path.join(repoRoot, ".github", "workflows", "windows-office-live.yml"), "utf8");
const windowsPptxLossless = await fs.readFile(path.join(repoRoot, ".github", "workflows", "windows-pptx-lossless.yml"), "utf8");

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
assert.match(fastSource, /compact-skill-jsons/);
assert.match(fastSource, /skill-json-assets/);
assert.match(fastSource, /windows-live-gate/);
assert.match(fastSource, /windows-pptx-lossless-gate/);
assert.doesNotMatch(fastSource, /default-template-library|agent-evals|pdf-provider-pack-build|pdf-pyhanko-provider|presentation-skill|document-skill|pdf-skill|reference-skills|office-kit-package/);
assert.match(slowSource, /default-template-library/);
assert.match(slowSource, /compact-skill-jsons/);
assert.match(slowSource, /skill-json-assets/);
assert.match(slowSource, /agent-evals/);
assert.match(slowSource, /pdf-provider-pack-build/);
const slowSegments = [
  "foundation",
  "presentation",
  "templates",
  "officekit",
  "documents",
  "pdf-packs",
  "pdf-providers",
  "pdf-specialists",
  "qa",
  "release",
];
const templateShards = ["documents-a", "documents-b", "presentations-a", "presentations-b", "spreadsheets-a", "spreadsheets-b"];
assert.match(slowSource, /const slowSegments = Object\.freeze/);
for (const shard of templateShards) assert.match(slowSource, new RegExp(`"--shard", "${shard}"`));
for (const segment of slowSegments) {
  assert.match(slowSource, new RegExp(`${segment.replaceAll("-", "\\-")}:?`));
  assert.match(slow, new RegExp(`npm run test:slow -- --segment ${segment}`));
  assert.match(release, new RegExp(`npm run test:slow -- --segment ${segment}`));
}
assert.doesNotMatch(slow, /run:\s*npm run test:slow\s*$/m);
assert.doesNotMatch(release, /run:\s*npm run test:slow\s*$/m);
assert.doesNotMatch(slow, /OFFICE_TEMPLATE_SOURCE_ROOT:/);
assert.doesNotMatch(release, /OFFICE_TEMPLATE_SOURCE_ROOT:/);

assert.match(ci, /name:\s*Fast gate/);
assert.match(ci, /npm test/);
assert.doesNotMatch(ci, /Install isolated PyMuPDF|default-template-library|agent-evals/);
assert.match(slow, /workflow_dispatch:/);
assert.match(slow, /schedule:/);
assert.match(slow, /cron:\s*"0 18 \* \* \*"/);
assert.doesNotMatch(slow, /^\s{2}(?:push|pull_request):/m);
assert.match(slow, /12-hour rolling cooldown/);
assert.match(slow, /npm run test:slow/);
assert.match(slow, /timeout-minutes:\s*45/);
assert.match(windows, /workflow_dispatch:/);
assert.match(windows, /self-hosted/);
assert.match(windows, /validate-windows-live-evidence/);
assert.doesNotMatch(windows, /test:excel-live|test:powerpoint-live/);
assert.match(windowsPptxLossless, /workflow_dispatch:/);
assert.match(windowsPptxLossless, /self-hosted/);
assert.match(windowsPptxLossless, /validate-windows-pptx-lossless-evidence/);
assert.match(windowsPptxLossless, /evidence_dir:/, "Windows PPTX lane must accept the complete evidence directory");
assert.match(windowsPptxLossless, /path: \$\{\{ inputs\.evidence_dir \}\}/, "Windows PPTX lane must upload rendered page evidence");
assert.match(windowsPptxLossless, /--verify-pixel-files/, "Windows PPTX lane must verify rendered image bytes");

console.log("gate policy contract ok");
