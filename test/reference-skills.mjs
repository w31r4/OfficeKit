import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

import {
  DocumentFile,
  FileBlob,
  PdfFile,
  PresentationFile,
  SpreadsheetFile,
  Workbook,
} from "../src/index.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const skillsRoot = path.join(repoRoot, "skills");
const pluginNames = [
  "documents",
  "spreadsheets",
  "presentations",
  "pdf",
  "office-kit",
  "template-creator",
  "presentation-template-creator",
  "default-template-library",
  "presentation-template-library",
  "skill-update",
];
const defaultTemplateSkills = [
  "artifact-template-analytics-dashboard",
  "artifact-template-design-report",
  "artifact-template-experiment-analysis",
  "artifact-template-financial-budget",
  "artifact-template-investment-committee-memo",
  "artifact-template-legal-memorandum",
  "artifact-template-minimal-letterhead",
  "artifact-template-operating-calendar",
  "artifact-template-project-tracker",
  "artifact-template-sales-pipeline",
  "artifact-template-strategy-memorandum",
  "artifact-template-system-design",
  "artifact-template-three-statement-forecast",
];
const presentationTemplateSkills = [
  "artifact-template-amber-committee-memo",
  "artifact-template-apricot-dossier",
  "artifact-template-aqua-impact-story",
  "artifact-template-axis-atlas",
  "artifact-template-blue-flame-operations",
  "artifact-template-blueprint-lecture",
  "artifact-template-business-review",
  "artifact-template-clay-craft-review",
  "artifact-template-coastal-analysis",
  "artifact-template-coral-growth-brief",
  "artifact-template-cranberry-evidence",
  "artifact-template-cream-civic-collage",
  "artifact-template-ebony-investment-review",
  "artifact-template-forest-strategy",
  "artifact-template-gilt-market-ledger",
  "artifact-template-grid-layout-library",
  "artifact-template-indigo-verdict",
  "artifact-template-jade-annual-brief",
  "artifact-template-lake-research-journal",
  "artifact-template-market-trends-report",
  "artifact-template-midnight-prospectus",
  "artifact-template-moonlit-work-report",
  "artifact-template-moss-transformation",
  "artifact-template-noir-field-pictorial",
  "artifact-template-operating-review",
  "artifact-template-paper-seminar",
  "artifact-template-project-kickoff",
  "artifact-template-rice-paper-yearbook",
  "artifact-template-river-handbook",
  "artifact-template-saffron-editorial",
  "artifact-template-silver-atelier",
  "artifact-template-simple-dark-mode",
  "artifact-template-simple-light-mode",
  "artifact-template-skyline-wayfinding",
  "artifact-template-soft-proof",
  "artifact-template-team-alignment",
  "artifact-template-tidal-research",
  "artifact-template-violet-operations",
];
const expectedSkills = new Map([
  ["documents", ["documents"]],
  ["spreadsheets", ["excel-live-control", "spreadsheets"]],
  ["presentations", ["powerpoint-live-control", "presentation-editorial-trim", "presentations"]],
  ["pdf", ["pdf"]],
  ["office-kit", ["office-kit"]],
  ["template-creator", ["template-creator"]],
  ["presentation-template-creator", ["presentation-template-creator"]],
  ["skill-update", ["skill-update"]],
  ["default-template-library", defaultTemplateSkills],
  ["presentation-template-library", presentationTemplateSkills],
]);
const expectedDeclaredSkillNames = new Map([
  ["documents", "documents"],
  ["excel-live-control", "excel-live-control"],
  ["spreadsheets", "Spreadsheets"],
  ["presentations", "Presentations"],
  ["presentation-editorial-trim", "presentation-editorial-trim"],
  ["powerpoint-live-control", "powerpoint-live-control"],
  ["pdf", "pdf"],
  ["office-kit", "office-kit"],
  ["template-creator", "template-creator"],
  ["presentation-template-creator", "presentation-template-creator"],
  ["skill-update", "skill-update"],
]);
for (const skillName of defaultTemplateSkills) expectedDeclaredSkillNames.set(skillName, skillName);
for (const skillName of presentationTemplateSkills) expectedDeclaredSkillNames.set(skillName, skillName);

async function exists(file) {
  return fs.access(file).then(() => true, () => false);
}

async function walk(root) {
  const files = [];
  for (const entry of await fs.readdir(root, { withFileTypes: true })) {
    const target = path.join(root, entry.name);
    if (entry.isDirectory()) files.push(...await walk(target));
    else if (entry.isFile()) files.push(target);
  }
  return files;
}

function yamlValue(source, key) {
  return source.match(new RegExp(`^\\s*${key}:\\s*["']?([^"'\\n]+)["']?\\s*$`, "m"))?.[1]?.trim();
}

for (const pluginName of pluginNames) {
  const pluginRoot = path.join(skillsRoot, pluginName);
  const manifestPath = path.join(pluginRoot, ".codex-plugin", "plugin.json");
  const manifest = JSON.parse(await fs.readFile(manifestPath, "utf8"));
  assert.equal(manifest.name, pluginName);
  const expectedVersion = new Set([
    "office-kit",
    "default-template-library",
    "presentation-template-creator",
    "presentation-template-library",
    "presentations",
    "template-creator",
    "skill-update",
  ]).has(pluginName) ? "1.1.0" : "0.2.0";
  assert.equal(manifest.version, expectedVersion);
  assert.equal(manifest.license, pluginName === "default-template-library" ? "MIT" : "AGPL-3.0-or-later");
  assert.equal(manifest.skills, "./skills/");
  assert.equal(manifest.repository, "https://github.com/w31r4/OfficeKit");
  assert.ok(await exists(path.join(pluginRoot, "README.md")));
  if (["documents", "spreadsheets", "presentations", "pdf"].includes(pluginName)) {
    const neutralManifest = JSON.parse(
      await fs.readFile(path.join(pluginRoot, "manifest.json"), "utf8"),
    );
    assert.equal(neutralManifest.schemaVersion, 1);
    assert.equal(neutralManifest.name, pluginName);
    assert.deepEqual(
      neutralManifest.skills,
      pluginName === "presentations"
        ? ["skills/presentations", "skills/presentation-editorial-trim", "skills/powerpoint-live-control"]
        : [`skills/${pluginName}`],
    );
    assert.ok(await exists(path.join(pluginRoot, neutralManifest.assets.icon)));
    assert.ok(await exists(path.join(pluginRoot, neutralManifest.assets.logo)));
    assert.ok(neutralManifest.interface?.displayName);
    const nativeAgent = await fs.readFile(
      path.join(pluginRoot, "skills", pluginName, "agents", "agent.yaml"),
      "utf8",
    );
    assert.equal(
      yamlValue(nativeAgent, "display_name"),
      neutralManifest.interface.displayName,
    );
    assert.ok(yamlValue(nativeAgent, "default_prompt"));
  }
  for (const iconKey of pluginName === "office-kit" ? [] : ["composerIcon", "logo"]) {
    assert.ok(await exists(path.resolve(pluginRoot, manifest.interface[iconKey])), `${pluginName} ${iconKey} must resolve inside the plugin`);
  }

  const skillNames = (await fs.readdir(path.join(pluginRoot, "skills"), { withFileTypes: true }))
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name)
    .sort();
  assert.deepEqual(skillNames, expectedSkills.get(pluginName));
  for (const skillName of skillNames) {
    const skillRoot = path.join(pluginRoot, "skills", skillName);
    const skillText = await fs.readFile(path.join(skillRoot, "SKILL.md"), "utf8");
    const frontmatter = skillText.match(/^---\n([\s\S]*?)\n---/);
    assert.ok(frontmatter, `${pluginName}/${skillName} is missing YAML frontmatter`);
    assert.equal(yamlValue(frontmatter[1], "name"), expectedDeclaredSkillNames.get(skillName));
    const templateSkill = ["default-template-library", "presentation-template-library"].includes(pluginName);
    const agentFilename = skillName.endsWith("template-creator") || templateSkill ? "agent.yaml" : "openai.yaml";
    const agentText = await fs.readFile(path.join(skillRoot, "agents", agentFilename), "utf8");
    for (const iconKey of pluginName === "office-kit" ? [] : templateSkill ? ["icon_large"] : ["icon_small", "icon_large"]) {
      const icon = yamlValue(agentText, iconKey);
      assert.ok(icon, `${pluginName}/${skillName} is missing ${iconKey}`);
      assert.ok(await exists(path.resolve(skillRoot, icon)), `${pluginName}/${skillName} ${iconKey} does not resolve`);
    }
  }
}

const officeKitRoot = path.join(skillsRoot, "office-kit", "skills", "office-kit");
const officeKitSkillText = await fs.readFile(path.join(officeKitRoot, "SKILL.md"), "utf8");
const officeKitRoutingText = await fs.readFile(path.join(officeKitRoot, "references", "routing.md"), "utf8");
const officeKitTemplateText = await fs.readFile(path.join(officeKitRoot, "references", "template-selection.md"), "utf8");
assert.match(officeKitSkillText, /exactly one owning Skill to each output/i);
assert.match(officeKitSkillText, /selected.*ask.*none/s);
assert.match(officeKitSkillText, /Do not use the Office template catalog for a PDF-only task/i);
assert.match(officeKitRoutingText, /One output has one owner/i);
assert.match(officeKitTemplateText, /`none` is a successful result/i);
assert.match(officeKitTemplateText, /original file stays in the task/i);
assert.match(officeKitTemplateText, /Search is local BM25F/i);
assert.match(officeKitTemplateText, /does not call a\s+model, build a vector/i);
assert.equal(
  await exists(path.join(officeKitRoot, "scripts", "query-templates.mjs")),
  false,
);
assert.ok(await exists(path.join(repoRoot, "src", "templates", "search.mjs")));

const templateCreatorManifest = JSON.parse(await fs.readFile(path.join(skillsRoot, "template-creator", "manifest.json"), "utf8"));
assert.equal(templateCreatorManifest.schemaVersion, 1);
assert.deepEqual(templateCreatorManifest.skills, ["skills/template-creator"]);
const presentationTemplateCreatorManifest = JSON.parse(await fs.readFile(path.join(skillsRoot, "presentation-template-creator", "manifest.json"), "utf8"));
assert.equal(presentationTemplateCreatorManifest.schemaVersion, 1);
assert.deepEqual(presentationTemplateCreatorManifest.skills, ["skills/presentation-template-creator"]);
const defaultTemplateManifest = JSON.parse(await fs.readFile(path.join(skillsRoot, "default-template-library", "manifest.json"), "utf8"));
assert.equal(defaultTemplateManifest.schemaVersion, 1);
assert.deepEqual(defaultTemplateManifest.skills, [
  "skills/artifact-template-design-report",
  "skills/artifact-template-experiment-analysis",
  "skills/artifact-template-investment-committee-memo",
  "skills/artifact-template-legal-memorandum",
  "skills/artifact-template-minimal-letterhead",
  "skills/artifact-template-strategy-memorandum",
  "skills/artifact-template-system-design",
  "skills/artifact-template-analytics-dashboard",
  "skills/artifact-template-financial-budget",
  "skills/artifact-template-operating-calendar",
  "skills/artifact-template-project-tracker",
  "skills/artifact-template-sales-pipeline",
  "skills/artifact-template-three-statement-forecast",
]);
const presentationTemplateManifest = JSON.parse(await fs.readFile(path.join(skillsRoot, "presentation-template-library", "manifest.json"), "utf8"));
assert.equal(presentationTemplateManifest.schemaVersion, 1);
assert.deepEqual(
  [...presentationTemplateManifest.skills].sort(),
  presentationTemplateSkills.map((skillName) => `skills/${skillName}`),
);
assert.equal(await exists(path.join(skillsRoot, "default-template-library", "LICENSE.md")), true);
assert.equal(await exists(path.join(skillsRoot, "default-template-library", "integrity.json")), true);
assert.equal(await exists(path.join(skillsRoot, "default-template-library", "catalog.json")), false);

assert.equal(await exists(path.join(skillsRoot, "documents", "SKILL.md")), false);
assert.equal(await exists(path.join(skillsRoot, "spreadsheets", "scripts")), false);
assert.equal(await exists(path.join(skillsRoot, "presentations", "fixtures")), false);
assert.ok(await exists(path.join(repoRoot, "test", "skill-harness", "spreadsheets", "scripts", "workflow.mjs")));
assert.ok(await exists(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "artifact_tool_docs", "API_QUICK_START.md")));
assert.ok(await exists(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "features", "charts.md")));
assert.ok(await exists(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "features", "pivot-tables.md")));
assert.equal(await exists(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "API_QUICK_START.md")), false);
assert.equal(await exists(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "charts.md")), false);
const spreadsheetSkillText = await fs.readFile(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "SKILL.md"), "utf8");
assert.match(spreadsheetSkillText, /artifact_tool_docs\/API_QUICK_START\.md/);
assert.match(spreadsheetSkillText, /features\/charts\.md/);
assert.match(spreadsheetSkillText, /features\/pivot-tables\.md/);
assert.match(spreadsheetSkillText, /officekit-pivot-table-workflow\.mjs/);
assert.match(spreadsheetSkillText, /officekit-financial-returns-workflow\.mjs/);
assert.match(spreadsheetSkillText, /officekit-loan-amortization-workflow\.mjs/);
assert.match(spreadsheetSkillText, /officekit-asset-depreciation-workflow\.mjs/);
assert.match(spreadsheetSkillText, /officekit-statistical-analysis-workflow\.mjs/);
assert.match(spreadsheetSkillText, /least-squares slope\/intercept\/R-squared\/standard-error.*LINEST.*FORECAST\.LINEAR.*TREND.*forecast sequence/i);
assert.match(spreadsheetSkillText, /officekit-exponential-growth-workflow\.mjs/);
assert.match(spreadsheetSkillText, /positive-y exponential model.*LOGEST.*GROWTH.*LN.*EXP/i);
assert.match(spreadsheetSkillText, /bounded `LET\(\.\.\.\)`.*_xlfn.*_xlws.*_xlpm.*ANCHORARRAY/i);
assert.match(spreadsheetSkillText, /officekit-growth-assumption-edit-workflow\.mjs/);
assert.match(spreadsheetSkillText, /officekit-connection-refresh-hardening-workflow\.mjs/);
assert.match(spreadsheetSkillText, /officekit-pivot-refresh-hardening-workflow\.mjs/);
assert.match(spreadsheetSkillText, /officekit-operating-plan-workflow\.mjs/);
assert.ok(await exists(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "examples", "officekit-growth-assumption-edit-workflow.mjs")));
assert.ok(await exists(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "examples", "officekit-connection-refresh-hardening-workflow.mjs")));
assert.ok(await exists(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "examples", "officekit-pivot-refresh-hardening-workflow.mjs")));
assert.ok(await exists(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "examples", "officekit-operating-plan-workflow.mjs")));
assert.ok(await exists(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "examples", "officekit-exponential-growth-workflow.mjs")));

const presentationApiRoot = path.join(skillsRoot, "presentations", "skills", "presentations", "artifact_tool", "api");
const presentationApiDocs = await fs.readFile(path.join(presentationApiRoot, "API_DOCS.md"), "utf8");
const presentationSpec = await fs.readFile(path.join(presentationApiRoot, "references", "presentation.spec.md"), "utf8");
const presentationSectionsSpec = await fs.readFile(path.join(presentationApiRoot, "references", "sections.spec.md"), "utf8");
const presentationLayoutSpec = await fs.readFile(path.join(presentationApiRoot, "references", "layout.spec.md"), "utf8");
const presentationChartSpec = await fs.readFile(path.join(presentationApiRoot, "references", "charts.spec.md"), "utf8");
assert.match(presentationApiDocs, /presentation\.view/);
assert.match(presentationApiDocs, /inkml-content-part-clone\.spec\.md/);
assert.match(presentationApiDocs, /references\/charts\.spec\.md/);
assert.match(presentationApiDocs, /sections\.spec\.md/);
assert.match(presentationSpec, /showGridlines\(\).*showGuides\(\)/s);
assert.match(presentationSpec, /gridSpacingCxEmu.*gridSpacingCyEmu/s);
assert.match(presentationSpec, /presentation\.view\.capability.*editable/s);
assert.match(presentationSpec, /setSourceProperties\(/);
assert.match(presentationSpec, /guide count\/order\/orientation|add\/remove\/reorient guides/i);
assert.match(presentationSpec, /presentation\.sections\.add/);
assert.match(presentationSectionsSpec, /p14:sectionLst/);
assert.match(presentationSectionsSpec, /partition every deck slide exactly once/i);
assert.match(presentationSectionsSpec, /officekit-section-rename-workflow\.mjs/);
assert.match(presentationSectionsSpec, /officekit-section-boundary-edit-workflow\.mjs/);
assert.match(presentationLayoutSpec, /read-only `slideGuides`/);
assert.match(presentationChartSpec, /standard `area`.*50%-hole `doughnut`/s);
assert.match(presentationChartSpec, /Marker-only `scatter`.*aligned.*`xValues`/s);
assert.match(presentationChartSpec, /2D `bubble`.*positive `bubbleSize`/s);
assert.match(presentationChartSpec, /formula references[\s\S]*fail\s+closed/i);
const presentationSkillRoot = path.join(skillsRoot, "presentations", "skills", "presentations");
const presentationSkillText = [
  await fs.readFile(path.join(presentationSkillRoot, "SKILL.md"), "utf8"),
  await fs.readFile(path.join(presentationSkillRoot, "references", "imported-capabilities.md"), "utf8"),
  await fs.readFile(path.join(presentationSkillRoot, "references", "source-continuation.md"), "utf8"),
].join("\n");
assert.match(presentationSkillText, /imported capabilities/i);
assert.match(presentationSkillText, /source continuation/i);
assert.match(presentationSkillText, /slide\.cloneCapability/);
assert.match(presentationSkillText, /presentation\.editNativeLeaf/);
assert.match(presentationSkillText, /officekit-chart-families-workflow\.mjs/);
for (const example of [
  "officekit-chart-families-workflow.mjs",
  "officekit-title-notes-edit-workflow.mjs",
  "officekit-legacy-comment-edit-workflow.mjs",
  "officekit-slide-name-edit-workflow.mjs",
  "officekit-view-properties-edit-workflow.mjs",
  "officekit-transition-edit-workflow.mjs",
  "officekit-section-rename-workflow.mjs",
  "officekit-section-boundary-edit-workflow.mjs",
]) {
  assert.ok(await exists(path.join(presentationSkillRoot, "examples", example)), example);
}

const documentsSkillRoot = path.join(skillsRoot, "documents", "skills", "documents");
const documentsManifest = (await fs.readFile(path.join(documentsSkillRoot, "manifest.txt"), "utf8"))
  .split(/\r?\n/)
  .map((entry) => entry.trim())
  .filter(Boolean);
assert.equal(new Set(documentsManifest).size, documentsManifest.length, "Documents manifest must not contain duplicates");
for (const entry of documentsManifest) {
  assert.equal(path.isAbsolute(entry), false, `Documents manifest entry must be relative: ${entry}`);
  assert.ok(!entry.split("/").includes(".."), `Documents manifest entry must stay inside the Skill: ${entry}`);
  assert.ok(await exists(path.join(documentsSkillRoot, entry)), `Documents manifest entry is missing: ${entry}`);
}
assert.ok(documentsManifest.includes("artifact_tool/API_QUICK_START.md"));
assert.ok(documentsManifest.includes("artifact_tool/ACCESSIBILITY_AUDIT.md"));
assert.ok(documentsManifest.includes("artifact_tool/_source_bound_docx.mjs"));
assert.ok(documentsManifest.includes("artifact_tool/_source_bound_sections.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-end-to-end.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-accessibility-audit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-classic-comment-edit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-board-review-surgical-edit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-image-alt-text-edit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-heading-level-edit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-hyperlink-text-edit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-section-page-numbering-edit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-section-margin-edit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-section-page-geometry-edit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-section-line-numbering-edit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-section-columns-edit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-section-break-edit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-table-column-widths-edit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-table-formatting-edit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-table-header-rows-edit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-table-row-break-policy-edit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-table-accessibility-edit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/officekit-note-text-edit-workflow.mjs"));
assert.ok(documentsManifest.includes("examples/end_to_end_smoke_test.md"));
assert.ok(await exists(path.join(documentsSkillRoot, "examples", "end_to_end_smoke_test.md")));
const documentsSkillText = await fs.readFile(path.join(documentsSkillRoot, "SKILL.md"), "utf8");
assert.match(documentsSkillText, /examples\/end_to_end_smoke_test\.md/);
assert.match(documentsSkillText, /officekit-section-page-numbering-edit-workflow\.mjs/);
assert.ok(await exists(path.join(documentsSkillRoot, "examples", "officekit-section-page-numbering-edit-workflow.mjs")));
assert.match(documentsSkillText, /officekit-image-alt-text-edit-workflow\.mjs/);
assert.ok(await exists(path.join(documentsSkillRoot, "examples", "officekit-image-alt-text-edit-workflow.mjs")));
assert.match(documentsSkillText, /officekit-heading-level-edit-workflow\.mjs/);
assert.ok(await exists(path.join(documentsSkillRoot, "examples", "officekit-heading-level-edit-workflow.mjs")));
assert.match(documentsSkillText, /officekit-hyperlink-text-edit-workflow\.mjs/);
assert.ok(await exists(path.join(documentsSkillRoot, "examples", "officekit-hyperlink-text-edit-workflow.mjs")));
assert.match(documentsSkillText, /officekit-section-margin-edit-workflow\.mjs/);
assert.ok(await exists(path.join(documentsSkillRoot, "examples", "officekit-section-margin-edit-workflow.mjs")));
assert.match(documentsSkillText, /officekit-section-page-geometry-edit-workflow\.mjs/);
assert.ok(await exists(path.join(documentsSkillRoot, "examples", "officekit-section-page-geometry-edit-workflow.mjs")));
assert.match(documentsSkillText, /officekit-section-line-numbering-edit-workflow\.mjs/);
assert.ok(await exists(path.join(documentsSkillRoot, "examples", "officekit-section-line-numbering-edit-workflow.mjs")));
assert.match(documentsSkillText, /officekit-section-columns-edit-workflow\.mjs/);
assert.ok(await exists(path.join(documentsSkillRoot, "examples", "officekit-section-columns-edit-workflow.mjs")));
assert.match(documentsSkillText, /officekit-section-break-edit-workflow\.mjs/);
assert.ok(await exists(path.join(documentsSkillRoot, "examples", "officekit-section-break-edit-workflow.mjs")));
assert.match(documentsSkillText, /officekit-table-column-widths-edit-workflow\.mjs/);
assert.ok(await exists(path.join(documentsSkillRoot, "examples", "officekit-table-column-widths-edit-workflow.mjs")));
assert.match(documentsSkillText, /officekit-table-formatting-edit-workflow\.mjs/);
assert.ok(await exists(path.join(documentsSkillRoot, "examples", "officekit-table-formatting-edit-workflow.mjs")));
assert.match(documentsSkillText, /officekit-table-header-rows-edit-workflow\.mjs/);
assert.ok(await exists(path.join(documentsSkillRoot, "examples", "officekit-table-header-rows-edit-workflow.mjs")));
assert.match(documentsSkillText, /officekit-table-row-break-policy-edit-workflow\.mjs/);
assert.ok(await exists(path.join(documentsSkillRoot, "examples", "officekit-table-row-break-policy-edit-workflow.mjs")));
assert.match(documentsSkillText, /officekit-table-accessibility-edit-workflow\.mjs/);
assert.ok(await exists(path.join(documentsSkillRoot, "examples", "officekit-table-accessibility-edit-workflow.mjs")));
assert.match(documentsSkillText, /officekit-note-text-edit-workflow\.mjs/);
assert.ok(await exists(path.join(documentsSkillRoot, "examples", "officekit-note-text-edit-workflow.mjs")));

const pdfSkillRoot = path.join(skillsRoot, "pdf", "skills", "pdf");
const pdfSkillText = await fs.readFile(path.join(pdfSkillRoot, "SKILL.md"), "utf8");
assert.match(pdfSkillText, /office-kit/);
assert.match(pdfSkillText, /PdfArtifact/);
assert.match(pdfSkillText, /createPdfjsParser/);
assert.match(pdfSkillText, /Poppler/);
assert.match(pdfSkillText, /ReportLab/);
assert.match(pdfSkillText, /pdfplumber/);
assert.match(pdfSkillText, /pypdf/);
assert.match(pdfSkillText, /PyMuPDF/);
assert.match(pdfSkillText, /pyHanko/);
assert.match(pdfSkillText, /veraPDF/);
assert.match(pdfSkillText, /rewrite/);
assert.match(pdfSkillText, /incremental/);
assert.match(pdfSkillText, /sanitize/);
assert.match(pdfSkillText, /silent fallback/i);
assert.match(pdfSkillText, /original bytes/i);
assert.ok(await exists(path.join(pdfSkillRoot, "artifact_tool", "API_QUICK_START.md")));
assert.ok(await exists(path.join(pdfSkillRoot, "examples", "public-api-end-to-end.mjs")));
assert.ok(await exists(path.join(pdfSkillRoot, "examples", "accessible-board-report.mjs")));
for (const relativePath of [
  "manifest.txt",
  "references/PROVIDER_MATRIX.md",
  "references/SAVE_POLICIES.md",
  "references/SECURITY_CHECKLIST.md",
  "references/PRODUCT_BOUNDARIES.md",
  "scripts/pdf_provider.py",
  "scripts/mupdf.mjs",
  "scripts/reportlab_create.py",
  "scripts/pdfplumber_extract.py",
  "scripts/pypdf_edit.py",
  "scripts/pymupdf_edit.py",
  "scripts/python_runtime.py",
  "scripts/residue_scan.py",
  "tasks/create.md",
  "tasks/read_review.md",
  "tasks/edit_existing.md",
  "tasks/forms_annotations.md",
  "tasks/sign_verify.md",
  "tasks/redact.md",
  "tasks/accessibility.md",
  "tasks/render_review.md",
]) assert.ok(await exists(path.join(pdfSkillRoot, relativePath)), `PDF Skill is missing ${relativePath}`);

const liveExcelSkillRoot = path.join(skillsRoot, "spreadsheets", "skills", "excel-live-control");
const liveExcelSkill = await fs.readFile(path.join(liveExcelSkillRoot, "SKILL.md"), "utf8");
const liveExcelProtocol = await fs.readFile(path.join(liveExcelSkillRoot, "references", "live-protocol.md"), "utf8");
assert.equal(await exists(path.join(skillsRoot, "spreadsheets", ".app.json")), false, "Excel Live Control must not retain a host connector declaration");
assert.match(liveExcelSkill, /officekit excel doctor --json/);
assert.match(liveExcelSkill, /officekit excel execute request\.json --json/);
assert.match(liveExcelSkill, /Home > Add-ins > My Add-ins > Upload My Add-in/);
assert.match(liveExcelSkill, /shared runtime/i);
assert.doesNotMatch(liveExcelSkill, /run_officejs|ChatGPT add-in|connected-document/i);
assert.match(liveExcelProtocol, /"protocol": 1/);
assert.match(liveExcelProtocol, /`pivot_table`/);
assert.match(liveExcelProtocol, /maybeApplied/);
assert.equal(await exists(path.join(liveExcelSkillRoot, "officejs.md")), false);

for (const file of (await walk(skillsRoot)).filter((item) => /\.(?:md|mjs|js|json|ya?ml|py)$/i.test(item))) {
  const source = await fs.readFile(file, "utf8");
  assert.doesNotMatch(source, /from\s+["']office-artifact-tool["']/, `${path.relative(repoRoot, file)} still imports the private package`);
}

const officialValidator = path.join(os.homedir(), ".codex", "skills", ".system", "plugin-creator", "scripts", "validate_plugin.py");
if (await exists(officialValidator)) {
  for (const pluginName of pluginNames.filter((name) => name !== "default-template-library")) {
    const validation = spawnSync("python3", [officialValidator, path.join(skillsRoot, pluginName)], { encoding: "utf8" });
    assert.equal(validation.status, 0, `${pluginName} failed the official plugin validator\n${validation.stdout}\n${validation.stderr}`);
  }
}

const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-reference-skills-"));
const previousPackageDir = process.env.OFFICE_KIT_PACKAGE_DIR;
try {
  process.env.OFFICE_KIT_PACKAGE_DIR = repoRoot;

  const { createDocument, DEFAULT_BRIEF } = await import(
    "../skills/documents/skills/documents/examples/officekit-end-to-end.mjs"
  );
  const docxPath = path.join(tempRoot, "officekit-decision-brief.docx");
  const authoredDocument = await createDocument(docxPath);
  assert.equal(authoredDocument.verification.ok, true);
  assert.match(authoredDocument.inspection.ndjson, /Launch readiness decision brief/);
  const documentRoundTrip = await DocumentFile.importDocx(await FileBlob.load(docxPath));
  assert.equal(documentRoundTrip.blocks.find((block) => block.kind === "table")?.getCell(1, 1).value, "Verified");
  assert.equal(documentRoundTrip.blocks.filter((block) => block.kind === "listItem").length, 3);
  assert.equal(documentRoundTrip.comments[0]?.text, "Recommendation wording verified for the release record.");
  assert.equal(documentRoundTrip.bookmarks[0]?.name, "DecisionSection");
  assert.deepEqual(documentRoundTrip.contentControls.map((control) => [control.tag, control.alias, control.controlType, control.controlType === "checkbox" ? control.checked : control.text]), [
    ["OWNER", "Brief owner", "text", DEFAULT_BRIEF.owner],
    ["FINAL_APPROVAL", "Final approval", "checkbox", true],
    ["REVIEW_PRIORITY", "Review priority", "dropdown", "High"],
    ["REVIEW_ROUTE", "Review route", "comboBox", "Security hotline"],
    ["REVIEW_DATE", "Review date", "date", "2026-07-21"],
  ]);
  assert.deepEqual(documentRoundTrip.notes.map((note) => [note.kind, note.text]), [
    ["footnote", "The final gate includes native rendering, package validation, and semantic re-import.\nThe delivery audit records the fixed physical paragraph count."],
    ["endnote", "Evidence snapshot dated 2026-07-17; retained with the release record."],
  ]);
  assert.deepEqual(documentRoundTrip.notes[0]?.paragraphs, [
    "The final gate includes native rendering, package validation, and semantic re-import.",
    "The delivery audit records the fixed physical paragraph count.",
  ]);
  assert.equal(documentRoundTrip.blocks.some(
    (block) => block.kind === "hyperlink" && block.anchor === "DecisionSection",
  ), true);
  assert.equal(documentRoundTrip.headers[0]?.text, "LAUNCH READINESS | DECISION BRIEF");
  assert.deepEqual(documentRoundTrip.footers[0]?.segments, [
    { text: "Page " },
    { field: { instruction: "PAGE", display: "1" } },
    { text: " of " },
    { field: { instruction: "NUMPAGES", display: "1" } },
  ]);
  assert.equal(documentRoundTrip.footers[0]?.fieldInstruction, "");
  assert.equal(documentRoundTrip.footers[0]?.editable, false);
  assert.deepEqual(documentRoundTrip.blocks.filter((block) => block.kind === "change").map(
    (block) => [block.changeType, block.text, block.author],
  ), [
    ["insert", "Final application-compatibility review is required before rollout.", "Lead reviewer"],
    ["delete", "Immediate unrestricted rollout.", "Release reviewer"],
  ]);
  const documentPackage = await JSZip.loadAsync(await fs.readFile(docxPath));
  const documentXml = await documentPackage.file("word/document.xml").async("text");
  assert.match(documentXml, /<w:ins\b/);
  assert.match(documentXml, /<w:del\b/);
  assert.match(documentXml, /<w:delText\b/);
  assert.match(documentXml, /<w:bookmarkStart\b[^>]*w:name="DecisionSection"/);
  assert.match(documentXml, /<w:bookmarkEnd\b/);
  assert.match(documentXml, /<w:hyperlink\b[^>]*w:anchor="DecisionSection"/);
  assert.match(documentXml, /<w:footnoteReference\b[^>]*w:id="1"/);
  assert.match(documentXml, /<w:endnoteReference\b[^>]*w:id="1"/);
  assert.match(documentXml, /<w:sdt>/);
  assert.match(documentXml, /<w:tag w:val="OWNER"\s*\/>/);
  assert.match(documentXml, /<w:tag w:val="FINAL_APPROVAL"\s*\/>/);
  assert.match(documentXml, /<w:tag w:val="REVIEW_ROUTE"\s*\/>[\s\S]*<w:comboBox w:lastValue="Security hotline">/);
  assert.match(documentXml, /<w:tag w:val="REVIEW_DATE"\s*\/>[\s\S]*<w:date w:fullDate="2026-07-21T00:00:00Z">/);
  assert.match(documentXml, /<w14:checkbox>[\s\S]*<w14:checked w14:val="1"\s*\/>/);
  assert.match(documentXml, /<w:t>Artifact Platform<\/w:t>/);
  const footnotesXml = await documentPackage.file("word/footnotes.xml").async("text");
  const endnotesXml = await documentPackage.file("word/endnotes.xml").async("text");
  for (const id of ["-1", "0", "1"]) assert.match(footnotesXml, new RegExp(`<w:footnote\\b[^>]*w:id="${id}"`));
  for (const id of ["-1", "0", "1"]) assert.match(endnotesXml, new RegExp(`<w:endnote\\b[^>]*w:id="${id}"`));
  assert.match(footnotesXml, /semantic re-import/);
  assert.match(endnotesXml, /retained with the release record/);

  const { ensureArtifactToolWorkspace, importArtifactTool } = await import(
    "../skills/presentations/skills/presentations/container_tools/artifact_tool_utils.mjs"
  );
  const workspace = path.join(tempRoot, "presentation-workspace");
  const prepared = await ensureArtifactToolWorkspace(workspace);
  assert.equal(prepared.packageDir, repoRoot);
  assert.equal(
    await fs.realpath(path.join(workspace, "node_modules", "office-kit")),
    await fs.realpath(repoRoot),
  );
  const importedPackage = await importArtifactTool(workspace);
  assert.equal(importedPackage.PresentationFile, PresentationFile);

  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("Summary");
  sheet.getRange("A1:C4").values = [
    ["Month", "Revenue", "EBITDA"],
    ["Jan", 100, 10],
    ["Feb", 120, 18],
    ["Mar", 130, 22],
  ];
  sheet.getRange("D1").values = [["Margin"]];
  sheet.getRange("D2").formulas = [["=C2/B2"]];
  sheet.getRange("D2:D4").fillDown();
  sheet.getRange("A1:D1").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" } };
  sheet.getRange("D2:D4").format.numberFormat = "0.0%";
  sheet.getRange("F1:G1").values = [["Month", "Revenue"]];
  sheet.getRange("F2:G2").formulas = [["=A2", "=B2"]];
  sheet.getRange("F2:G4").fillDown();
  const chart = sheet.charts.add("line", sheet.getRange("F1:G4"));
  chart.title = "Revenue Trend";
  chart.hasLegend = false;
  chart.setPosition("I1", "P15");
  const preview = await workbook.render({ sheetName: "Summary", autoCrop: "all", format: "svg" });
  assert.equal(preview.type, "image/svg+xml");
  assert.match(await preview.text(), /Revenue Trend/);
  const xlsx = await SpreadsheetFile.exportXlsx(workbook);
  const workbookRoundTrip = await SpreadsheetFile.importXlsx(xlsx);
  assert.deepEqual(workbookRoundTrip.worksheets.getItem("Summary").getRange("D2:D4").formulas, [
    ["=C2/B2"],
    ["=C3/B3"],
    ["=C4/B4"],
  ]);
  const csvWorkbook = await Workbook.fromCSV("Name,Value\nOfficeKit,1", { sheetName: "Data" });
  assert.deepEqual(csvWorkbook.worksheets.getItem("Data").getRange("A1:B2").values, [["Name", "Value"], ["OfficeKit", "1"]]);

  const { createWorkbook: createReferenceWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/officekit-range-workflow.mjs"
  );
  const spreadsheetPath = path.join(tempRoot, "officekit-range-workflow.xlsx");
  const authoredWorkbook = await createReferenceWorkbook(spreadsheetPath);
  assert.equal(authoredWorkbook.verification.ok, true);
  assert.match(authoredWorkbook.inspection.ndjson, /Revenue trend/);
  const spreadsheetRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(spreadsheetPath));
  assert.equal(spreadsheetRoundTrip.worksheets.getItem("Forecast").getRange("D3").format.numberFormat, "0.00%");
  assert.equal(spreadsheetRoundTrip.worksheets.getItem("Forecast").getRange("B3").formulasR1C1[0][0], "=R[-1]C*(1+'Assumptions'!R2C2)");

  const { createSparklineWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/officekit-sparkline-workflow.mjs"
  );
  const sparklinePath = path.join(tempRoot, "officekit-sparkline-workflow.xlsx");
  const authoredSparklines = await createSparklineWorkbook(sparklinePath);
  assert.equal(authoredSparklines.verification.ok, true);
  assert.match(authoredSparklines.inspection.ndjson, /"kind":"sparkline"/);
  const sparklineRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(sparklinePath));
  assert.deepEqual(sparklineRoundTrip.worksheets.getItem("Operating Trends").sparklineGroups.items.map((group) => group.type), ["line", "column"]);
  assert.equal(sparklineRoundTrip.worksheets.getItem("Operating Trends").sparklineGroups.items[0].seriesColor, "#F97316");

  const { createDataTableWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/officekit-data-table-workflow.mjs"
  );
  const dataTablePath = path.join(tempRoot, "officekit-data-table-workflow.xlsx");
  const authoredDataTables = await createDataTableWorkbook(dataTablePath);
  assert.equal(authoredDataTables.verification.ok, true);
  assert.match(authoredDataTables.inspection.ndjson, /"kind":"dataTable"/);
  const dataTableRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(dataTablePath));
  assert.deepEqual(
    dataTableRoundTrip.worksheets.getItem("Scenario Analysis").dataTables.__getDefinitions().map((item) => item.displayFormula),
    ["{=TABLE(D1)}", "{=TABLE(D1,D2)}"],
  );

  const { createPivotTableWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/officekit-pivot-table-workflow.mjs"
  );
  const pivotTablePath = path.join(tempRoot, "officekit-pivot-table-workflow.xlsx");
  const authoredPivotTable = await createPivotTableWorkbook(pivotTablePath);
  assert.equal(authoredPivotTable.verification.ok, true);
  assert.match(authoredPivotTable.inspection.ndjson, /"kind":"pivotTable"/);
  const pivotTableRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(pivotTablePath));
  const pivotTable = pivotTableRoundTrip.worksheets.getItem("Pivot Summary").pivotTables.items[0];
  assert.equal(pivotTable.name, "Revenue and units by region");
  assert.deepEqual(pivotTable.rowFields, ["Region", "Channel"]);
  assert.deepEqual(pivotTable.filters, [{ field: "Region", exclude: ["North"] }]);
  assert.deepEqual(pivotTable.computedValues().at(-1), ["Grand Total", "", 260, 25, 180, 19, 440, 44]);

  const { createFinancialReturnsWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/officekit-financial-returns-workflow.mjs"
  );
  const financialReturnsPath = path.join(tempRoot, "officekit-financial-returns-workflow.xlsx");
  const authoredFinancialReturns = await createFinancialReturnsWorkbook(financialReturnsPath);
  assert.equal(authoredFinancialReturns.verification.ok, true);
  assert.match(authoredFinancialReturns.inspection.ndjson, /XIRR/);
  assert.match(authoredFinancialReturns.inspection.ndjson, /MIRR/);
  const financialReturnsRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(financialReturnsPath));
  financialReturnsRoundTrip.recalculate();
  assert.equal(financialReturnsRoundTrip.worksheets.getItem("Returns").getRange("B8").formulas[0][0], "=XIRR('Inputs'!$C$14:$C$18,'Inputs'!$B$14:$B$18,'Inputs'!$B$7)");
  assert.equal(financialReturnsRoundTrip.worksheets.getItem("Returns").getRange("B9").formulas[0][0], "=MIRR('Inputs'!$C$14:$C$18,'Inputs'!$B$5,'Inputs'!$B$6)");
  assert.ok(Math.abs(financialReturnsRoundTrip.worksheets.getItem("Returns").getRange("B9").values[0][0] - 0.14400168352963139) < 1e-9);
  assert.deepEqual(financialReturnsRoundTrip.worksheets.getItem("Checks").getRange("E4:E10").values, [["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"]]);

  const { createLoanAmortizationWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/officekit-loan-amortization-workflow.mjs"
  );
  const loanAmortizationPath = path.join(tempRoot, "officekit-loan-amortization-workflow.xlsx");
  const authoredLoanAmortization = await createLoanAmortizationWorkbook(loanAmortizationPath);
  assert.equal(authoredLoanAmortization.verification.ok, true);
  assert.match(authoredLoanAmortization.inspection.ndjson, /PPMT/);
  assert.match(authoredLoanAmortization.checksInspection.ndjson, /PV/);
  assert.match(authoredLoanAmortization.checksInspection.ndjson, /FV/);
  assert.match(authoredLoanAmortization.checksInspection.ndjson, /NPER/);
  assert.match(authoredLoanAmortization.checksInspection.ndjson, /CUMIPMT/);
  assert.match(authoredLoanAmortization.checksInspection.ndjson, /CUMPRINC/);
  const loanAmortizationRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(loanAmortizationPath));
  loanAmortizationRoundTrip.recalculate();
  assert.equal(loanAmortizationRoundTrip.worksheets.getItem("Amortization").getRange("D5").formulas[0][0], "=IPMT('Inputs'!$B$10,A5,'Inputs'!$B$11,'Inputs'!$B$5,0,'Inputs'!$B$9)");
  assert.ok(Math.abs(loanAmortizationRoundTrip.worksheets.getItem("Amortization").getRange("F16").values[0][0]) < 1e-7);
  assert.equal(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B9").formulas[0][0], "=RATE('Inputs'!$B$11,'Amortization'!$C$5,'Inputs'!$B$5,0,'Inputs'!$B$9,'Inputs'!$B$10)");
  assert.ok(Math.abs(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B9").values[0][0] - 0.01) < 1e-10);
  assert.equal(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B10").formulas[0][0], "=PV('Inputs'!$B$10,'Inputs'!$B$11,'Amortization'!$C$5,0,'Inputs'!$B$9)");
  assert.equal(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B11").formulas[0][0], "=FV('Inputs'!$B$10,'Inputs'!$B$11,'Amortization'!$C$5,'Inputs'!$B$5,'Inputs'!$B$9)");
  assert.equal(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B12").formulas[0][0], "=NPER('Inputs'!$B$10,'Amortization'!$C$5,'Inputs'!$B$5,0,'Inputs'!$B$9)");
  assert.equal(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B13").formulas[0][0], "=CUMIPMT('Inputs'!$B$10,'Inputs'!$B$11,'Inputs'!$B$5,1,'Inputs'!$B$11,'Inputs'!$B$9)");
  assert.equal(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B14").formulas[0][0], "=CUMPRINC('Inputs'!$B$10,'Inputs'!$B$11,'Inputs'!$B$5,1,'Inputs'!$B$11,'Inputs'!$B$9)");
  assert.ok(Math.abs(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B10").values[0][0] - 100000) < 1e-7);
  assert.ok(Math.abs(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B11").values[0][0]) < 1e-7);
  assert.ok(Math.abs(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B12").values[0][0] - 12) < 1e-10);
  assert.ok(Math.abs(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B13").values[0][0] + 6618.54641401005) < 1e-8);
  assert.ok(Math.abs(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B14").values[0][0] + 100000) < 1e-8);
  assert.deepEqual(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("E4:E15").values, Array.from({ length: 12 }, () => ["OK"]));

  const { createAssetDepreciationWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/officekit-asset-depreciation-workflow.mjs"
  );
  const assetDepreciationPath = path.join(tempRoot, "officekit-asset-depreciation-workflow.xlsx");
  const authoredAssetDepreciation = await createAssetDepreciationWorkbook(assetDepreciationPath);
  assert.equal(authoredAssetDepreciation.verification.ok, true);
  assert.match(authoredAssetDepreciation.inspection.ndjson, /SLN/);
  assert.match(authoredAssetDepreciation.inspection.ndjson, /DDB/);
  assert.match(authoredAssetDepreciation.inspection.ndjson, /SYD/);
  const assetDepreciationRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(assetDepreciationPath));
  assetDepreciationRoundTrip.recalculate();
  assert.equal(assetDepreciationRoundTrip.worksheets.getItem("Depreciation").getRange("D5").formulas[0][0], "=DB('Inputs'!$B$5,'Inputs'!$B$6,'Inputs'!$B$7,A5,'Inputs'!$B$8)");
  assert.equal(assetDepreciationRoundTrip.worksheets.getItem("Depreciation").getRange("E5").formulas[0][0], "=DDB('Inputs'!$B$5,'Inputs'!$B$6,'Inputs'!$B$7,A5,'Inputs'!$B$9)");
  assert.equal(assetDepreciationRoundTrip.worksheets.getItem("Depreciation").getRange("H5").formulas[0][0], "=SYD('Inputs'!$B$5,'Inputs'!$B$6,'Inputs'!$B$7,A5)");
  assert.equal(assetDepreciationRoundTrip.worksheets.getItem("Depreciation").getRange("H5").values[0][0], 30000);
  assert.equal(assetDepreciationRoundTrip.worksheets.getItem("Depreciation").getRange("F9").values[0][0], 10000);
  assert.deepEqual(assetDepreciationRoundTrip.worksheets.getItem("Checks").getRange("E4:E10").values, Array.from({ length: 7 }, () => ["OK"]));

  const { createStatisticalAnalysisWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/officekit-statistical-analysis-workflow.mjs"
  );
  const statisticalAnalysisPath = path.join(tempRoot, "officekit-statistical-analysis-workflow.xlsx");
  const authoredStatisticalAnalysis = await createStatisticalAnalysisWorkbook(statisticalAnalysisPath);
  assert.equal(authoredStatisticalAnalysis.verification.ok, true);
  assert.match(authoredStatisticalAnalysis.inspection.ndjson, /CORREL/);
  assert.match(authoredStatisticalAnalysis.inspection.ndjson, /COVARIANCE\.S/);
  assert.match(authoredStatisticalAnalysis.inspection.ndjson, /FORECAST\.LINEAR/);
  assert.match(authoredStatisticalAnalysis.inspection.ndjson, /LINEST/);
  assert.match(authoredStatisticalAnalysis.inspection.ndjson, /TREND/);
  const statisticalAnalysisRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(statisticalAnalysisPath));
  statisticalAnalysisRoundTrip.recalculate();
  assert.equal(statisticalAnalysisRoundTrip.worksheets.getItem("Analysis").getRange("B11").formulas[0][0], "=CORREL('Data'!$B$4:$B$9,'Data'!$C$4:$C$9)");
  assert.ok(Math.abs(statisticalAnalysisRoundTrip.worksheets.getItem("Analysis").getRange("B12").values[0][0] - 686) < 1e-9);
  assert.equal(statisticalAnalysisRoundTrip.worksheets.getItem("Analysis").getRange("B21").formulas[0][0], "=FORECAST.LINEAR(B20,'Data'!$C$4:$C$9,'Data'!$B$4:$B$9)");
  assert.ok(Math.abs(statisticalAnalysisRoundTrip.worksheets.getItem("Analysis").getRange("B21").values[0][0] - 138.6) < 1e-9);
  assert.equal(statisticalAnalysisRoundTrip.worksheets.getItem("Analysis").getRange("E16").formulas[0][0], "=LINEST('Data'!$C$4:$C$9,'Data'!$B$4:$B$9,TRUE,TRUE)");
  assert.equal(statisticalAnalysisRoundTrip.worksheets.getItem("Analysis").store.get("E16").dynamicArrayRef, "E16:F20");
  assert.ok(Math.abs(statisticalAnalysisRoundTrip.worksheets.getItem("Analysis").getRange("E16").values[0][0] - 1.96) < 1e-9);
  assert.equal(statisticalAnalysisRoundTrip.worksheets.getItem("Analysis").getRange("I16").formulas[0][0], "=TREND('Data'!$C$4:$C$9,'Data'!$B$4:$B$9,H16:H18)");
  assert.equal(statisticalAnalysisRoundTrip.worksheets.getItem("Analysis").store.get("I16").dynamicArrayRef, "I16:I18");
  assert.deepEqual(statisticalAnalysisRoundTrip.worksheets.getItem("Analysis").getRange("I16:I18").values, [[138.6], [158.2], [177.8]]);
  assert.deepEqual(statisticalAnalysisRoundTrip.worksheets.getItem("Checks").getRange("E4:E18").values, Array.from({ length: 15 }, () => ["OK"]));

  const { createExponentialGrowthWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/officekit-exponential-growth-workflow.mjs"
  );
  const exponentialGrowthPath = path.join(tempRoot, "officekit-exponential-growth-workflow.xlsx");
  const authoredExponentialGrowth = await createExponentialGrowthWorkbook(exponentialGrowthPath);
  assert.equal(authoredExponentialGrowth.verification.ok, true);
  assert.match(authoredExponentialGrowth.inspection.ndjson, /LOGEST/);
  assert.match(authoredExponentialGrowth.inspection.ndjson, /GROWTH/);
  const exponentialGrowthRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(exponentialGrowthPath));
  exponentialGrowthRoundTrip.recalculate();
  assert.equal(exponentialGrowthRoundTrip.worksheets.getItem("Analysis").getRange("E4").formulas[0][0], "=LOGEST('Data'!$B$4:$B$9,'Data'!$A$4:$A$9,TRUE,TRUE)");
  assert.equal(exponentialGrowthRoundTrip.worksheets.getItem("Analysis").store.get("E4").dynamicArrayRef, "E4:F8");
  assert.ok(Math.abs(exponentialGrowthRoundTrip.worksheets.getItem("Analysis").getRange("E4").values[0][0] - 1.71968316255041) < 1e-9);
  assert.equal(exponentialGrowthRoundTrip.worksheets.getItem("Analysis").getRange("I4").formulas[0][0], "=GROWTH('Data'!$B$4:$B$9,'Data'!$A$4:$A$9,H4:H6)");
  assert.equal(exponentialGrowthRoundTrip.worksheets.getItem("Analysis").store.get("I4").dynamicArrayRef, "I4:I6");
  assert.ok(Math.abs(exponentialGrowthRoundTrip.worksheets.getItem("Analysis").getRange("I6").values[0][0] - 473.980713611783) < 1e-9);
  assert.deepEqual(exponentialGrowthRoundTrip.worksheets.getItem("Checks").getRange("E4:E11").values, Array.from({ length: 8 }, () => ["OK"]));

  const { createRobustStatisticsWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/officekit-robust-statistics-workflow.mjs"
  );
  const robustStatisticsPath = path.join(tempRoot, "officekit-robust-statistics-workflow.xlsx");
  const authoredRobustStatistics = await createRobustStatisticsWorkbook(robustStatisticsPath);
  assert.equal(authoredRobustStatistics.verification.ok, true);
  assert.match(authoredRobustStatistics.inspection.ndjson, /RANK\.AVG/);
  assert.match(authoredRobustStatistics.inspection.ndjson, /MODE\.MULT/);
  const robustStatisticsRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(robustStatisticsPath));
  robustStatisticsRoundTrip.recalculate();
  assert.deepEqual(robustStatisticsRoundTrip.worksheets.getItem("Analysis").getRange("B4:B8").values.flat(), [5.5, 3, 2, 2.8, 2]);
  assert.equal(robustStatisticsRoundTrip.worksheets.getItem("Analysis").store.get("D4").dynamicArrayRef, "D4:D5");
  assert.deepEqual(robustStatisticsRoundTrip.worksheets.getItem("Analysis").getRange("D4:E5").values, [[2, 2], [3, 2]]);
  assert.deepEqual(robustStatisticsRoundTrip.worksheets.getItem("Checks").getRange("E4:E10").values, Array.from({ length: 7 }, () => ["OK"]));

  const { createScatterWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/officekit-scatter-chart-workflow.mjs"
  );
  const scatterPath = path.join(tempRoot, "officekit-scatter-chart-workflow.xlsx");
  const authoredScatter = await createScatterWorkbook(scatterPath);
  assert.equal(authoredScatter.verification.ok, true);
  assert.match(authoredScatter.inspection.ndjson, /"chartType":"scatter"/);
  const scatterRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(scatterPath));
  const scatterChart = scatterRoundTrip.worksheets.getItem("Relationship Analysis").charts.items[0];
  assert.equal(scatterChart.type, "scatter");
  assert.deepEqual(scatterChart.series.items[0].xValues, [10, 20, 25, 34, 45]);

  const { createBubbleWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/officekit-bubble-chart-workflow.mjs"
  );
  const bubblePath = path.join(tempRoot, "officekit-bubble-chart-workflow.xlsx");
  const authoredBubble = await createBubbleWorkbook(bubblePath);
  assert.equal(authoredBubble.verification.ok, true);
  assert.match(authoredBubble.inspection.ndjson, /"chartType":"bubble"/);
  const bubbleRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(bubblePath));
  const bubbleChart = bubbleRoundTrip.worksheets.getItem("Opportunity Analysis").charts.items[0];
  assert.equal(bubbleChart.type, "bubble");
  assert.deepEqual(bubbleChart.series.items[0].bubbleSizes, [4, 10, 12, 18, 27]);

  const { createPdf } = await import(
    "../skills/pdf/skills/pdf/examples/public-api-end-to-end.mjs"
  );
  const pdfPath = path.join(tempRoot, "release-readiness-scorecard.pdf");
  const pdfRenderDir = path.join(tempRoot, "release-readiness-scorecard-pages");
  const authoredPdf = await createPdf(pdfPath, { renderDir: pdfRenderDir });
  assert.equal(authoredPdf.verification.ok, true);
  assert.equal(authoredPdf.fileInspection.summary.tagged, true);
  assert.equal(authoredPdf.fileInspection.summary.figures, 2);
  assert.equal(authoredPdf.renderedPages.length, authoredPdf.pdf.pages.length);
  assert.ok(authoredPdf.renderedPages.every((page) => page.bytes > 1_000));
  const pdfRoundTrip = await PdfFile.importPdf(await FileBlob.load(pdfPath));
  assert.equal(pdfRoundTrip.pages[0].tables[0].getCell(3, 2).value, "Verified");
  assert.match(pdfRoundTrip.extractText(), /Release readiness scorecard/);
  const { createPdfjsParser } = await import("office-kit/pdf/pdfjs");
  const parsedPdf = await PdfFile.importPdf(await FileBlob.load(pdfPath), {
    parser: createPdfjsParser(),
    preferParser: true,
    parserName: "pdfjs",
  });
  assert.match(parsedPdf.extractText(), /Release readiness scorecard/);
} finally {
  if (previousPackageDir === undefined) delete process.env.OFFICE_KIT_PACKAGE_DIR;
  else process.env.OFFICE_KIT_PACKAGE_DIR = previousPackageDir;
  await fs.rm(tempRoot, { recursive: true, force: true });
}

console.log("reference skill plugins smoke ok");
