#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const mode = process.argv[2] || "fast";
const gateArgs = process.argv.slice(3);

let segment;
if (gateArgs.length > 0) {
  if (gateArgs.length !== 2 || gateArgs[0] !== "--segment" || !gateArgs[1]) {
    console.error(`invalid gate options ${JSON.stringify(gateArgs)}; expected --segment <name> for the slow gate`);
    process.exit(2);
  }
  segment = gateArgs[1];
}

const nodeStep = (script, ...args) => ({
  label: `node ${script}${args.length ? ` ${args.join(" ")}` : ""}`,
  command: process.execPath,
  args: [script, ...args],
});

const npmStep = (script) => ({
  label: `npm run ${script}`,
  command: process.platform === "win32" ? "npm.cmd" : "npm",
  args: ["run", script],
  shell: process.platform === "win32",
});

const fastSteps = [
  nodeStep("scripts/check-js-syntax.mjs"),
  nodeStep("scripts/optimize-skill-pngs.mjs", "--check"),
  nodeStep("test/skill-png-assets.mjs"),
  nodeStep("scripts/compact-skill-jsons.mjs", "--check"),
  nodeStep("test/skill-json-assets.mjs"),
  nodeStep("test/skill-portability.mjs"),
  nodeStep("test/office-kit.mjs"),
  nodeStep("test/officekit-identity.mjs"),
  nodeStep("test/font-metrics.mjs"),
  nodeStep("test/ooxml-source-reference.mjs"),
  nodeStep("test/spreadsheet.mjs"),
  nodeStep("test/spreadsheet-statistics.mjs"),
  nodeStep("test/spreadsheet-formula-syntax.mjs"),
  nodeStep("test/spreadsheet-data-table.mjs"),
  nodeStep("test/spreadsheet-range-compat.mjs"),
  nodeStep("test/spreadsheet-sparkline.mjs"),
  nodeStep("test/presentation.mjs"),
  nodeStep("test/pptx-lossless-benchmark.mjs"),
  nodeStep("test/pptx-design-profile.mjs"),
  nodeStep("test/pptx-source-continuation.mjs"),
  nodeStep("test/presentation-jsx.mjs"),
  nodeStep("test/document.mjs"),
  nodeStep("test/pdf.mjs"),
  nodeStep("test/office-kit-skill.mjs"),
  nodeStep("test/officekit-cli.mjs"),
  npmStep("test:repl"),
  nodeStep("test/reference-skill-sync.mjs"),
  nodeStep("test/claude-plugin.mjs"),
  nodeStep("test/gate-policy.mjs"),
  nodeStep("test/windows-live-gate.mjs"),
  nodeStep("test/windows-pptx-lossless-gate.mjs"),
  nodeStep("test/verify.mjs"),
  nodeStep("test/review.mjs"),
  nodeStep("test/help.mjs"),
  nodeStep("test/package-contents.mjs"),
];

const slowSteps = [
  nodeStep("scripts/optimize-skill-pngs.mjs", "--check"),
  nodeStep("test/skill-png-assets.mjs"),
  nodeStep("scripts/compact-skill-jsons.mjs", "--check"),
  nodeStep("test/skill-json-assets.mjs"),
  nodeStep("test/skill-portability.mjs"),
  nodeStep("test/office-kit.mjs"),
  nodeStep("test/officekit-identity.mjs"),
  nodeStep("test/font-metrics.mjs"),
  nodeStep("test/ooxml-source-reference.mjs"),
  nodeStep("test/spreadsheet.mjs"),
  nodeStep("test/spreadsheet-statistics.mjs"),
  nodeStep("test/spreadsheet-formula-syntax.mjs"),
  nodeStep("test/spreadsheet-data-table.mjs"),
  nodeStep("test/spreadsheet-range-compat.mjs"),
  nodeStep("test/spreadsheet-sparkline.mjs"),
  npmStep("test:skill:spreadsheets"),
  nodeStep("test/presentation.mjs"),
  nodeStep("test/pptx-lossless-benchmark.mjs"),
  nodeStep("test/pptx-source-continuation-native.mjs"),
  nodeStep("test/pptx-source-agent-continuation.mjs"),
  nodeStep("test/presentation-custom-geometry-render.mjs"),
  nodeStep("test/presentation-jsx.mjs"),
  npmStep("test:skill:presentations"),
  nodeStep("test/template-library.mjs"),
  nodeStep("test/default-template-library.mjs", "--shard", "documents-a"),
  nodeStep("test/default-template-library.mjs", "--shard", "documents-b"),
  nodeStep("test/default-template-library.mjs", "--shard", "presentations-a"),
  nodeStep("test/default-template-library.mjs", "--shard", "presentations-b"),
  nodeStep("test/default-template-library.mjs", "--shard", "spreadsheets-a"),
  nodeStep("test/default-template-library.mjs", "--shard", "spreadsheets-b"),
  nodeStep("test/template-creator.mjs"),
  nodeStep("test/office-kit-skill.mjs"),
  nodeStep("test/officekit-cli.mjs"),
  nodeStep("test/officekit-tasks.mjs"),
  nodeStep("test/officekit-repl.mjs"),
  nodeStep("test/officekit-repl-interrupted-write.mjs"),
  nodeStep("test/officekit-task-artifacts.mjs"),
  npmStep("test:excel-live"),
  npmStep("test:powerpoint-live"),
  nodeStep("test/document.mjs"),
  nodeStep("test/document-skill.mjs"),
  nodeStep("test/document-table-formatting-workflow.mjs"),
  nodeStep("test/document-table-header-rows-workflow.mjs"),
  nodeStep("test/document-table-row-break-policy-workflow.mjs"),
  nodeStep("test/document-table-accessibility-workflow.mjs"),
  nodeStep("test/document-image-alt-text-workflow.mjs"),
  nodeStep("test/pdf.mjs"),
  nodeStep("test/pdf-providers.mjs"),
  nodeStep("test/pdf-provider-pack-build.mjs"),
  nodeStep("test/pdf-provider-release-tools.mjs"),
  nodeStep("test/pdf-python-pack-build.mjs"),
  nodeStep("test/pdf-ocr-pack-build.mjs"),
  nodeStep("test/pdf-poppler-pack-build.mjs"),
  nodeStep("test/pdf-verapdf-pack-build.mjs"),
  nodeStep("test/pdf-qpdf-managed-release.mjs"),
  nodeStep("test/pdf-python-managed-release.mjs"),
  nodeStep("test/pdf-ocr-managed-release.mjs"),
  nodeStep("test/pdf-verapdf-managed-release.mjs"),
  nodeStep("test/pdf-poppler-managed-release.mjs"),
  nodeStep("test/pdf-skill.mjs"),
  nodeStep("test/pdf-pypdf-flatten.mjs"),
  nodeStep("test/pdf-provider-skill.mjs"),
  nodeStep("test/pdf-qpdf-provider.mjs"),
  nodeStep("test/pdf-pikepdf-provider.mjs"),
  nodeStep("test/pdf-pyhanko-provider.mjs"),
  nodeStep("test/pdf-pyhanko-certified-form-fill.mjs"),
  nodeStep("test/pdf-verapdf-provider.mjs"),
  nodeStep("test/pdf-ocrmypdf-provider.mjs"),
  nodeStep("test/reference-skill-sync.mjs"),
  nodeStep("test/reference-skills.mjs"),
  nodeStep("test/claude-plugin.mjs"),
  nodeStep("test/agent-evals.mjs"),
  nodeStep("test/agent-eval-branded-template.mjs"),
  nodeStep("test/verify.mjs"),
  nodeStep("test/review.mjs"),
  nodeStep("test/render.mjs"),
  nodeStep("test/visual-baselines.mjs"),
  nodeStep("test/renderer-adapters.mjs"),
  nodeStep("test/playwright-renderer.mjs"),
  nodeStep("test/office-bridge.mjs"),
  nodeStep("test/examples.mjs"),
  nodeStep("test/release-check.mjs"),
  nodeStep("test/package-contents.mjs"),
  nodeStep("test/standalone-distribution.mjs"),
  nodeStep("test/help.mjs"),
];

const slowSegments = Object.freeze({
  foundation: Object.freeze({ start: 0, end: 16 }),
  presentation: Object.freeze({ start: 16, end: 23 }),
  templates: Object.freeze({ start: 23, end: 30 }),
  officekit: Object.freeze({ start: 30, end: 37 }),
  documents: Object.freeze({ start: 37, end: 46 }),
  "pdf-packs": Object.freeze({ start: 46, end: 55 }),
  "pdf-providers": Object.freeze({ start: 55, end: 64 }),
  "pdf-specialists": Object.freeze({ start: 64, end: 73 }),
  qa: Object.freeze({ start: 73, end: 80 }),
  release: Object.freeze({ start: 80, end: 85 }),
});

const segmentRanges = Object.values(slowSegments);
if (
  segmentRanges[0]?.start !== 0
  || segmentRanges.at(-1)?.end !== slowSteps.length
  || segmentRanges.some((range, index) => index > 0 && range.start !== segmentRanges[index - 1].end)
) {
  throw new Error("slow gate segments must form one contiguous cover of slowSteps");
}

if (segment && mode !== "slow") {
  console.error(`test gate segments are only available for slow; received ${JSON.stringify(mode)}`);
  process.exit(2);
}

const selectedSlowRange = segment ? slowSegments[segment] : null;
if (segment && !selectedSlowRange) {
  console.error(`unknown slow gate segment ${JSON.stringify(segment)}; expected one of ${Object.keys(slowSegments).join(", ")}`);
  process.exit(2);
}

const steps = mode === "fast"
  ? fastSteps
  : mode === "slow"
    ? selectedSlowRange
      ? slowSteps.slice(selectedSlowRange.start, selectedSlowRange.end)
      : slowSteps
    : null;
if (!steps) {
  console.error(`unknown test gate ${JSON.stringify(mode)}; expected fast or slow`);
  process.exit(2);
}

const gateLabel = segment ? `${mode}/${segment}` : mode;
for (const [index, step] of steps.entries()) {
  console.error(`[${gateLabel} gate ${index + 1}/${steps.length}] ${step.label}`);
  const result = spawnSync(step.command, step.args, {
    cwd: repoRoot,
    env: process.env,
    stdio: "inherit",
    shell: step.shell ?? false,
  });
  if (result.error) {
    console.error(`[${gateLabel} gate] failed to start ${step.label}: ${result.error.message}`);
    process.exit(1);
  }
  if (result.status !== 0) {
    if (result.signal) console.error(`[${gateLabel} gate] ${step.label} terminated by ${result.signal}`);
    process.exit(result.status ?? 1);
  }
}

console.error(`${gateLabel} gate passed (${steps.length} steps)`);
