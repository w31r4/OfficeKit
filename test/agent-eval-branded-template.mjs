import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { editBrandedTemplate } from "../skills/presentations/skills/presentations/examples/officekit-branded-template-local-update-workflow.mjs";
import { loadSuite } from "../scripts/run-agent-evals.mjs";
import { gradeBrandedTemplateCase } from "../scripts/agent-eval-branded-template-grader.mjs";
import { verifiedLockedAsset } from "../scripts/run-agent-evals.mjs";

const { cases } = await loadSuite();
const item = cases.find((candidate) => candidate.id === "pptx-branded-template-local-update");
assert.equal(item?.status, "ready");
const sourcePath = await verifiedLockedAsset("presentations/quarterly-board-template.pptx");
const imagePath = await verifiedLockedAsset("presentations/replacement-product.png");
const root = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-branded-template-test-"));
const workspace = path.join(root, "workspace");
await fs.mkdir(path.join(workspace, "inputs"), { recursive: true });
await fs.mkdir(path.join(workspace, "outputs"), { recursive: true });
await fs.copyFile(sourcePath, path.join(workspace, "inputs", "template.pptx"));
await fs.copyFile(imagePath, path.join(workspace, "inputs", "replacement-product.png"));
await editBrandedTemplate({
  inputPath: path.join(workspace, "inputs", "template.pptx"),
  imagePath: path.join(workspace, "inputs", "replacement-product.png"),
  outputPath: path.join(workspace, "outputs", "quarterly-board-updated.pptx"),
  auditPath: path.join(workspace, "outputs", "audit.json"),
});
const trace = JSON.stringify({ type: "item.completed", item: { id: "branded", type: "command_execution", command: "node officekit-branded-template-local-update-workflow.mjs" } });
const report = await gradeBrandedTemplateCase({ item, workspace, finalMessage: "typed OfficeKit branded-template transaction completed; provider.silentFallback=false", trace });
assert.equal(report.supported, true);
assert.equal(report.graded, true);
assert.equal(report.caseSpecificPassed, true, JSON.stringify(report.checks, null, 2));
assert.equal(report.rawScorePercent, 100);
console.log("Agent branded-template asset, typed workflow, preservation oracle, and render smoke passed.");
