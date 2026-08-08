import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const packageMetadata = JSON.parse(await fs.readFile(path.join(repoRoot, "package.json"), "utf8"));
assert.equal(packageMetadata.version, "0.6.0");
assert.equal(packageMetadata.license, "AGPL-3.0-or-later");
assert.equal(packageMetadata.dependencies.mupdf, "1.28.0");
assert.equal(packageMetadata.dependencies["@firecrawl/anydoc"], "0.1.3");
assert.equal(packageMetadata.dependencies.selfsigned, "^5.5.0");
assert.equal(packageMetadata.exports["./pdf/mupdf"], "./src/pdf/mupdf.mjs");
assert.equal(packageMetadata.exports["./pdf/providers"], "./src/pdf/providers/index.mjs");
assert.equal(packageMetadata.exports["./codec"], "./src/codecs/office-kit.mjs");
assert.equal(
  packageMetadata.exports["./codec/wire"],
  "./src/generated/office_kit/artifact/v1/office_artifact_pb.js",
);
assert.equal(packageMetadata.exports["./codecs/office-kit"], undefined);
assert.equal(packageMetadata.exports["./codecs/openxml-wasm"], undefined);
assert.deepEqual(packageMetadata.bin, {
  officekit: "./bin/officekit.mjs",
});
assert.equal(packageMetadata.engines.node, ">=22.15.0");
assert.equal(
  packageMetadata.scripts["build:standalone"],
  "node scripts/build-standalone.mjs",
);
assert.equal(
  packageMetadata.scripts["build:excel-addin"],
  "node scripts/build-excel-addin.mjs",
);
assert.equal(
  packageMetadata.scripts["test:excel-live"],
  "node scripts/build-excel-addin.mjs && node test/excel-live.mjs",
);
assert.equal(packageMetadata.scripts["build:powerpoint-addin"], "node scripts/build-powerpoint-addin.mjs");
assert.equal(packageMetadata.scripts["test:powerpoint-live"], "node scripts/build-powerpoint-addin.mjs && node test/powerpoint-live.mjs");
assert.equal(packageMetadata.exports["./live"], "./src/live/bridge.mjs");
assert.equal(packageMetadata.exports["./live/protocol"], "./src/live/protocol.mjs");
assert.equal(packageMetadata.exports["./live/adapters/powerpoint"], "./src/live/adapters/powerpoint.mjs");
assert.equal(packageMetadata.exports["./powerpoint-live"], "./src/powerpoint-live/repl.mjs");
assert.equal(
  packageMetadata.scripts["test:standalone"],
  "node test/standalone-distribution.mjs",
);
assert.equal(packageMetadata.scripts.postinstall, undefined, "MuPDF must not require npm lifecycle hooks");
const nodeRuntimes = JSON.parse(
  await fs.readFile(path.join(repoRoot, "standalone", "node-runtimes.v1.json"), "utf8"),
);
assert.equal(nodeRuntimes.schemaVersion, 1);
assert.equal(nodeRuntimes.nodeVersion, "24.18.0");
assert.deepEqual(Object.keys(nodeRuntimes.runtimes).sort(), [
  "darwin-arm64",
  "linux-x64",
  "win32-x64",
]);
for (const [target, runtime] of Object.entries(nodeRuntimes.runtimes)) {
  assert.ok(["darwin-arm64", "linux-x64", "win32-x64"].includes(target));
  assert.match(runtime.url, /^https:\/\/nodejs\.org\/dist\/v24\.18\.0\//);
  assert.match(runtime.sha256, /^[a-f0-9]{64}$/);
  assert.ok(Number.isSafeInteger(runtime.size) && runtime.size > 30_000_000);
}
const standaloneReleases = JSON.parse(
  await fs.readFile(path.join(repoRoot, "standalone", "releases.v1.json"), "utf8"),
);
assert.equal(standaloneReleases.officeKitVersion, packageMetadata.version);
assert.deepEqual(Object.keys(standaloneReleases.assets).sort(), [
  "darwin-arm64",
  "linux-x64",
  "win32-x64",
]);
for (const [target, release] of Object.entries(standaloneReleases.assets)) {
  const extension = target === "win32-x64" ? ".zip" : ".tar.gz";
  assert.equal(release.asset, `office-kit-${packageMetadata.version}-${target}${extension}`);
  assert.match(release.sha256, /^[a-f0-9]{64}$/);
  assert.ok(Number.isSafeInteger(release.size) && release.size > 30_000_000);
}
const standaloneInstaller = await fs.readFile(
  path.join(repoRoot, "standalone", "install.sh"),
  "utf8",
);
assert.match(standaloneInstaller, /OFFICE_KIT_VERSION=0\.6\.0/);
assert.doesNotMatch(standaloneInstaller, /FINALIZE_/);
const windowsStandaloneInstaller = await fs.readFile(
  path.join(repoRoot, "standalone", "install.ps1"),
  "utf8",
);
assert.match(windowsStandaloneInstaller, /\$OfficeKitVersion = "0\.6\.0"/);
assert.match(windowsStandaloneInstaller, /win32-x64/);
assert.doesNotMatch(windowsStandaloneInstaller, /RELEASE_(?:SHA256|SIZE)/);
const pdfFacadeSource = await fs.readFile(path.join(repoRoot, "src", "pdf", "index.mjs"), "utf8");
assert.match(pdfFacadeSource, /await import\("\.\/mupdf\.mjs"\)/, "MuPDF must load only when a PDF operation needs it");
assert.doesNotMatch(pdfFacadeSource, /from\s+["']mupdf["']/, "the root PDF facade must not initialize MuPDF eagerly");
const pdfProvidersSource = await fs.readFile(path.join(repoRoot, "src", "pdf", "providers", "index.mjs"), "utf8");
assert.doesNotMatch(pdfProvidersSource, /from\s+["']mupdf["']/, "the explicit provider subpath must not initialize MuPDF eagerly");
const reviewSource = await fs.readFile(path.join(repoRoot, "src", "review", "index.mjs"), "utf8");
assert.match(reviewSource, /await import\("@firecrawl\/anydoc"\)/, "AnyDoc must load only for an explicitly requested content view");
assert.doesNotMatch(reviewSource, /from\s+["']@firecrawl\/anydoc["']/, "the review facade must not initialize AnyDoc eagerly");
const officeKitCliSource = await fs.readFile(path.join(repoRoot, "src", "cli", "officekit.mjs"), "utf8");
assert.doesNotMatch(officeKitCliSource, /node:child_process|https?:\/\/|\bfetch\s*\(/, "officekit init must remain a local Skill installer");
assert.doesNotMatch(officeKitCliSource, /pdf\/providers|from\s+["']mupdf["']/, "officekit init must not initialize PDF runtimes or capability packs");
assert.match(officeKitCliSource, /await import\("\.\.\/excel-live\/cli\.mjs"\)/, "Excel Live Control must load only for the excel subcommand");
assert.match(officeKitCliSource, /await import\("\.\/repl\.mjs"\)/, "REPL must load only for the repl subcommand");
assert.doesNotMatch(officeKitCliSource, /from\s+["']\.\.\/excel-live\//, "root CLI import must not start the Excel bridge");
const templateSearchSource = await fs.readFile(path.join(repoRoot, "src", "templates", "search.mjs"), "utf8");
assert.doesNotMatch(templateSearchSource, /pdf\/providers|from\s+["']mupdf["']|runtime\/office-kit/, "template search must not initialize Office or PDF runtimes");
const presentationCodecSource = await fs.readFile(path.join(repoRoot, "src", "codecs", "office-kit-presentation.mjs"), "utf8");
assert.match(presentationCodecSource, /from "\.\.\/presentation\/index\.mjs";/, "the Presentation codec must depend on the Presentation leaf module");
assert.match(presentationCodecSource, /from "\.\/office-kit-presentation-charts\.mjs";/, "the Presentation codec must delegate chart wire semantics to the chart leaf module");
assert.doesNotMatch(presentationCodecSource, /from "\.\.\/index\.mjs";/, "the Presentation codec must not create a back-edge to the root entry");
const presentationFacadeSource = await fs.readFile(path.join(repoRoot, "src", "presentation", "index.mjs"), "utf8");
const presentationCodecAdapterSource = await fs.readFile(path.join(repoRoot, "src", "codecs", "office-kit-presentation-codec.mjs"), "utf8");
const officeKitRuntimeSource = await fs.readFile(path.join(repoRoot, "src", "codecs", "office-kit-runtime.mjs"), "utf8");
assert.match(presentationFacadeSource, /await import\("\.\.\/codecs\/office-kit-presentation-codec\.mjs"\)/, "Presentation file I/O must load the format-specific codec adapter");
assert.doesNotMatch(presentationFacadeSource, /await import\("\.\.\/codecs\/office-kit\.mjs"\)/, "Presentation file I/O must not load the aggregate Document/Spreadsheet codec");
assert.match(presentationCodecAdapterSource, /from "\.\/office-kit-runtime\.mjs";/, "the Presentation adapter must use the shared runtime boundary");
assert.match(presentationCodecAdapterSource, /from "\.\/office-kit-presentation\.mjs";/, "the Presentation adapter must use the Presentation wire mapper");
assert.doesNotMatch(presentationCodecAdapterSource, /from "\.\/office-kit\.mjs";/, "the Presentation adapter must not load the aggregate codec");
assert.doesNotMatch(presentationCodecAdapterSource, /\.\.\/(?:document|spreadsheet)\/index\.mjs/, "the Presentation adapter must not load another artifact model");
assert.doesNotMatch(officeKitRuntimeSource, /from "\.\/office-kit\.mjs";/, "the runtime boundary must not load the aggregate codec");
assert.doesNotMatch(officeKitRuntimeSource, /\.\.\/(?:document|presentation|spreadsheet)\/index\.mjs/, "the runtime boundary must not own an artifact model");
const presentationChartCodecSource = await fs.readFile(path.join(repoRoot, "src", "codecs", "office-kit-presentation-charts.mjs"), "utf8");
assert.doesNotMatch(presentationChartCodecSource, /from "\.\.\/index\.mjs";/, "the Presentation chart codec must not create a back-edge to the root entry");
const spreadsheetCodecSource = await fs.readFile(path.join(repoRoot, "src", "codecs", "office-kit.mjs"), "utf8");
assert.match(spreadsheetCodecSource, /from "\.\.\/spreadsheet\/index\.mjs";/, "the Spreadsheet codec must depend on the Spreadsheet leaf module");
assert.doesNotMatch(spreadsheetCodecSource, /from "\.\.\/index\.mjs";/, "the Spreadsheet codec must not create a back-edge to the root entry");
const spreadsheetLeafSource = await fs.readFile(path.join(repoRoot, "src", "spreadsheet", "index.mjs"), "utf8");
const formulaEngineSource = await fs.readFile(path.join(repoRoot, "src", "spreadsheet", "formula-engine.mjs"), "utf8");
assert.match(spreadsheetLeafSource, /from "\.\/formula-engine\.mjs";/, "the Spreadsheet leaf must own the formula-engine dependency boundary");
assert.doesNotMatch(formulaEngineSource, /from\s+["']\.\/index\.mjs["']/, "the formula engine must operate on workbook data shape without a Spreadsheet-model back-edge");
assert.doesNotMatch(formulaEngineSource, /\bFunction\s*\(/, "the bounded formula evaluator must not execute generated JavaScript");
const skillsNpmIgnore = await fs.readFile(path.join(repoRoot, "skills", ".npmignore"), "utf8");
assert.match(skillsNpmIgnore, /__pycache__/);
assert.match(skillsNpmIgnore, /\*\.pyc/);
const result = spawnSync("npm", ["pack", "--dry-run", "--json", "--ignore-scripts"], {
  cwd: repoRoot,
  encoding: "utf8",
});
assert.equal(result.status, 0, `npm pack manifest failed\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`);
const report = JSON.parse(result.stdout)[0];
const files = report.files.map((item) => item.path);
// npm's gzip output varies between the macOS and Linux npm builds used by local
// and hosted gates. The 0.6.0 global CLI deliberately ships the twenty audited
// default DOCX/XLSX/PPTX templates once inside the package. Keep narrow
// cross-platform headroom over the measured 36,175,810-byte archive.
const maxPackedBytes = 37_500_000;
// The bundled OfficeKit runtime is an audited product payload, not an
// optional download. Keep its unpacked budget tight while allowing the
// audited PDF provider/docs growth plus the bounded Office codecs and runnable
// workflows. The managed-capability resolver distributes only catalog, policy,
// and installer source -- never specialist binaries. Keep bounded headroom for
// its Skill/API contract without concealing a runtime bundle in the npm tarball.
// The MIT Default Template Library is a deliberate consumer payload: init
// references it in place and never copies it into a project. Its twenty
// retained Office files, previews, metadata cards, and Skill instructions
// account for the 0.4.0 budget increase. PowerPoint sections plus the bounded transition and rich
// speaker-notes leaves, the public formula catalog, bounded formula expression
// parser, SUMPRODUCT range-mask profile, source-bound DOCX header/footer,
// source-bound PowerPoint section-name and complete-boundary transactions, and
// source-free structured page-furniture fields and source-bound PPTX
// view-properties mutation,
// the compact OfficeKit routing Skill plus its local BM25F template retrieval,
// XLSX connection-refresh and imported-Pivot refresh-on-load transactions,
// formula-input syntax guard, and source-bound embedded-DOCX OLE package
// replacement, plus the canonical DOCX 1-through-16-paragraph note body and
// source-bound section line-numbering, column-profile, break-type, fixed-table
// column-width, direct-formatting, repeat-header-row, and image-alt-text
// transactions add protobuf, audited WASM, public Help, and native guidance;
// retain measured headroom instead of hiding that product surface. The
// Source-aware formula introspection and OOXML input-budget Help/API evidence
// add a small measured payload; the bounded ceiling moves by 10 KiB rather
// than silently allowing an unbounded package increase.
const maxUnpackedBytes = 53_520_000;
// Public Skill PNGs are required user-facing assets. They are retained with
// byte-identical non-IDAT chunks and inflated scanline streams, but their IDAT
// payloads are deterministically recompressed. Prevent future PNG tooling from
// silently consuming the recovered product-growth headroom.
const maxSkillPngBytes = 3_550_000;

for (const required of [
  "LICENSE",
  ".claude-plugin/marketplace.json",
  "apps/excel-addin/dist/taskpane.html",
  "apps/excel-addin/dist/taskpane.js",
  "apps/excel-addin/dist/taskpane.css",
  "apps/excel-addin/dist/support.html",
  "apps/excel-addin/dist/assets/officekit-excel-32.png",
  "apps/excel-addin/dist/assets/officekit-excel-80.png",
  "apps/powerpoint-addin/dist/taskpane.html",
  "apps/powerpoint-addin/dist/taskpane.js",
  "apps/powerpoint-addin/dist/taskpane.css",
  "apps/powerpoint-addin/dist/support.html",
  "apps/powerpoint-addin/dist/assets/officekit-powerpoint-32.png",
  "apps/powerpoint-addin/dist/assets/officekit-powerpoint-80.png",
  "README.md",
  "README.zh-CN.md",
  "THIRD_PARTY_NOTICES.md",
  "bin/officekit.mjs",
  "docs/api.md",
  "docs/reference-skills.md",
  "docs/template-library-provenance.md",
  "proto/office_kit/artifact/v1/office_artifact.proto",
  "src/generated/office_kit/artifact/v1/office_artifact_pb.js",
  "src/codecs/office-kit.mjs",
  "src/codecs/office-kit-error.mjs",
  "src/codecs/office-kit-runtime.mjs",
  "src/codecs/office-kit-assets.mjs",
  "src/codecs/office-kit-presentation-codec.mjs",
  "src/codecs/office-kit-presentation.mjs",
  "src/codecs/office-kit-presentation-charts.mjs",
  "src/codecs/office-kit-spreadsheet-pivots.mjs",
  "runtime/office-kit/main.mjs",
  "runtime/office-kit/manifest.json",
  "runtime/office-kit/sbom.cdx.json",
  "runtime/office-kit/DOTNET-LICENSE.TXT",
  "runtime/office-kit/DOTNET-THIRD-PARTY-NOTICES.TXT",
  "runtime/office-kit/_framework/dotnet.native.wasm",
  "runtime/office-kit/_framework/OfficeKit.Codec.wasm",
  "runtime/office-kit/_framework/OfficeKit.Runtime.wasm",
  "src/ooxml/docx-comments.mjs",
  "src/ooxml/docx-bibliography.mjs",
  "src/ooxml/package.mjs",
  "src/presentation/chart-trendline-svg.mjs",
  "src/presentation/index.mjs",
  "src/presentation/ooxml-chart-data.mjs",
  "src/presentation/ooxml-charts.mjs",
  "src/presentation/ooxml-hyperlinks.mjs",
  "src/presentation/ooxml-custom-shows.mjs",
  "src/ooxml/docx-links.mjs",
  "src/ooxml/docx-numbering.mjs",
  "src/ooxml/docx-sections.mjs",
  "src/pdf/table-grid.mjs",
  "src/pdf/reading-order.mjs",
  "src/pdf/accessibility.mjs",
  "src/pdf/index.mjs",
  "src/pdf/mupdf.mjs",
  "src/review/index.mjs",
  "skills/office-kit/skills/office-kit/references/review.md",
  "src/pdf/providers/catalog.mjs",
  "src/pdf/providers/index.mjs",
  "src/pdf/providers/installer.mjs",
  "src/pdf/providers/policy.mjs",
  "src/pdf/providers/provider-catalog.v1.json",
  "src/document/index.mjs",
  "src/cli/officekit.mjs",
  "src/cli/run-task.mjs",
  "src/cli/officekit-resolver.mjs",
  "src/cli/repl.mjs",
  "src/excel-live/bridge.mjs",
  "src/excel-live/bridge-server.mjs",
  "src/excel-live/certificates.mjs",
  "src/excel-live/cli.mjs",
  "src/excel-live/client.mjs",
  "src/excel-live/errors.mjs",
  "src/excel-live/manifest.mjs",
  "src/excel-live/protocol.mjs",
  "src/excel-live/state.mjs",
  "src/excel-live/repl.mjs",
  "src/live/bridge.mjs",
  "src/live/errors.mjs",
  "src/live/protocol.mjs",
  "src/live/adapters/index.mjs",
  "src/live/adapters/excel.mjs",
  "src/live/adapters/powerpoint.mjs",
  "src/live/cli.mjs",
  "src/powerpoint-live/bridge-server.mjs",
  "src/powerpoint-live/client.mjs",
  "src/powerpoint-live/manifest.mjs",
  "src/powerpoint-live/repl.mjs",
  "src/powerpoint-live/state.mjs",
  "src/templates/search.mjs",
  "src/help/index.mjs",
  "src/index.mjs",
  "examples/officekit-repl-cells.jsonl",
  "src/ooxml/docx-source-references.mjs",
  "src/ooxml/docx-settings.mjs",
  "src/ooxml/pptx-package-semantics.mjs",
  "src/ooxml/pptx-source-references.mjs",
  "src/ooxml/source-reference-xml.mjs",
  "src/ooxml/source-references.mjs",
  "src/presentation/ooxml-theme.mjs",
  "src/presentation/group-shapes.mjs",
  "src/presentation/native-objects.mjs",
  "src/presentation/compose.mjs",
  "src/presentation/custom-geometry.mjs",
  "src/presentation/text-paragraphs.mjs",
  "src/presentation/ooxml-masters.mjs",
  "src/presentation/ooxml-modern-comments.mjs",
  "src/shared/colors.mjs",
  "src/shared/binary.mjs",
  "src/shared/file-blob.mjs",
  "src/shared/ids.mjs",
  "src/shared/images.mjs",
  "src/shared/inspection.mjs",
  "src/shared/png.mjs",
  "src/shared/render-output.mjs",
  "src/shared/text-range.mjs",
  "src/shared/xml.mjs",
  "src/spreadsheet/formula-criteria.mjs",
  "src/spreadsheet/formula-engine.mjs",
  "src/spreadsheet/index.mjs",
  "src/spreadsheet/data-validations.mjs",
  "src/spreadsheet/data-tables.mjs",
  "src/codecs/office-kit-spreadsheet-data-tables.mjs",
  "src/spreadsheet/formula-coercion.mjs",
  "src/spreadsheet/chart-source-data.mjs",
  "src/spreadsheet/ooxml-styles.mjs",
  "src/spreadsheet/ooxml-threaded-comments.mjs",
  "src/spreadsheet/pivot-dates.mjs",
  "src/spreadsheet/pivot-filters.mjs",
  "src/spreadsheet/pivot-formulas.mjs",
  "src/spreadsheet/pivot-groups.mjs",
  "src/spreadsheet/pivots.mjs",
  "src/spreadsheet/range-addressing.mjs",
  "src/spreadsheet/range-operations.mjs",
  "src/spreadsheet/structured-references.mjs",
  "src/spreadsheet/worksheet-protection.mjs",
  "src/codecs/office-kit-spreadsheet-protection.mjs",
  "native/OfficeBridge/src/OfficeBridge.csproj",
  "skills/documents/.codex-plugin/plugin.json",
  "skills/documents/manifest.json",
  "skills/documents/README.md",
  "skills/documents/assets/icon.png",
  "skills/documents/skills/documents/SKILL.md",
  "skills/documents/skills/documents/agents/openai.yaml",
  "skills/documents/skills/documents/agents/agent.yaml",
  "skills/documents/skills/documents/LICENSE.txt",
  "skills/documents/skills/documents/artifact_tool/API_QUICK_START.md",
  "skills/documents/skills/documents/artifact_tool/_source_bound_docx.mjs",
  "skills/documents/skills/documents/artifact_tool/_source_bound_sections.mjs",
  "skills/documents/skills/documents/examples/officekit-end-to-end.mjs",
  "skills/documents/skills/documents/examples/officekit-classic-comment-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/officekit-board-review-surgical-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/officekit-page-furniture-text-edit.mjs",
  "skills/documents/skills/documents/examples/officekit-header-text-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/officekit-footer-text-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/officekit-image-alt-text-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/officekit-section-page-numbering-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/officekit-section-margin-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/officekit-section-page-geometry-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/officekit-section-line-numbering-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/officekit-section-columns-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/officekit-section-break-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/officekit-table-column-widths-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/officekit-table-formatting-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/officekit-table-header-rows-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/officekit-table-row-break-policy-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/officekit-table-accessibility-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/officekit-note-text-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/end_to_end_smoke_test.md",
  "skills/documents/skills/documents/render_docx.py",
  "skills/documents/skills/documents/scripts/docx_ooxml_patch.py",
  "skills/documents/skills/documents/tasks/create_edit.md",
  "skills/spreadsheets/.codex-plugin/plugin.json",
  "skills/spreadsheets/manifest.json",
  "skills/spreadsheets/README.md",
  "skills/spreadsheets/skills/spreadsheets/SKILL.md",
  "skills/spreadsheets/skills/spreadsheets/agents/openai.yaml",
  "skills/spreadsheets/skills/spreadsheets/agents/agent.yaml",
  "skills/spreadsheets/skills/spreadsheets/artifact_tool_docs/API_QUICK_START.md",
  "skills/spreadsheets/skills/spreadsheets/features/charts.md",
  "skills/spreadsheets/skills/spreadsheets/features/pivot-tables.md",
  "skills/spreadsheets/skills/spreadsheets/examples/officekit-range-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/officekit-sparkline-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/officekit-data-table-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/officekit-data-validation-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/officekit-worksheet-protection-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/officekit-pivot-table-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/officekit-loan-amortization-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/officekit-asset-depreciation-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/officekit-scatter-chart-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/officekit-bubble-chart-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/officekit-growth-assumption-edit-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/officekit-connection-refresh-hardening-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/officekit-pivot-refresh-hardening-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/officekit-opaque-enterprise-local-edit-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/officekit-operating-plan-workflow.mjs",
  "skills/spreadsheets/skills/excel-live-control/SKILL.md",
  "skills/spreadsheets/skills/excel-live-control/agents/openai.yaml",
  "skills/spreadsheets/skills/excel-live-control/assets/file-spreadsheet.png",
  "skills/spreadsheets/skills/excel-live-control/references/live-protocol.md",
  "skills/presentations/.codex-plugin/plugin.json",
  "skills/presentations/manifest.json",
  "skills/presentations/README.md",
  "skills/presentations/skills/presentations/SKILL.md",
  "skills/presentations/skills/presentations/agents/openai.yaml",
  "skills/presentations/skills/presentations/agents/agent.yaml",
  "skills/presentations/skills/powerpoint-live-control/SKILL.md",
  "skills/presentations/skills/powerpoint-live-control/agents/openai.yaml",
  "skills/presentations/skills/powerpoint-live-control/assets/icon.svg",
  "skills/presentations/skills/powerpoint-live-control/references/live-protocol.md",
  "skills/presentations/skills/presentations/style_guidelines.md",
  "skills/presentations/skills/presentations/routing/google_slides.md",
  "skills/presentations/skills/presentations/artifact_tool/API_QUICK_START.md",
  "skills/presentations/skills/presentations/examples/officekit-chart-families-workflow.mjs",
  "skills/presentations/skills/presentations/examples/officekit-legacy-comment-add-workflow.mjs",
  "skills/presentations/skills/presentations/examples/officekit-legacy-comment-edit-workflow.mjs",
  "skills/presentations/skills/presentations/examples/officekit-speaker-notes-add-workflow.mjs",
  "skills/presentations/skills/presentations/examples/officekit-title-notes-edit-workflow.mjs",
  "skills/presentations/skills/presentations/examples/officekit-rich-speaker-notes-edit-workflow.mjs",
  "skills/presentations/skills/presentations/examples/officekit-slide-name-edit-workflow.mjs",
  "skills/presentations/skills/presentations/examples/officekit-view-properties-edit-workflow.mjs",
  "skills/presentations/skills/presentations/examples/officekit-transition-edit-workflow.mjs",
  "skills/presentations/skills/presentations/examples/officekit-section-rename-workflow.mjs",
  "skills/presentations/skills/presentations/examples/officekit-section-boundary-edit-workflow.mjs",
  "skills/presentations/skills/presentations/examples/officekit-custom-show-workflow.mjs",
  "skills/presentations/skills/presentations/examples/officekit-slide-duplicate-workflow.mjs",
  "skills/presentations/skills/presentations/examples/officekit-smartart-text-edit-workflow.mjs",
  "skills/presentations/skills/presentations/examples/officekit-ole-office-package-workflow.mjs",
  "skills/presentations/skills/presentations/artifact_tool/api/references/custom-shows.spec.md",
  "skills/presentations/skills/presentations/artifact_tool/api/references/sections.spec.md",
  "skills/presentations/skills/presentations/artifact_tool/api/references/transitions.spec.md",
  "skills/presentations/skills/presentations/artifact_tool/api/references/ole-workbooks.spec.md",
  "skills/presentations/skills/presentations/artifact_tool/api/references/smartart-clone.spec.md",
  "skills/presentations/skills/presentations/artifact_tool/api/references/inkml-content-part-clone.spec.md",
  "skills/presentations/skills/presentations/artifact_tool/api/references/embedded-video-clone.spec.md",
  "skills/presentations/skills/presentations/container_tools/artifact_tool_utils.mjs",
  "skills/presentations/skills/presentations/container_tools/slides_test.py",
  "skills/presentations/skills/presentations/builtin_templates_support/scripts/create-presentation.mjs",
  "skills/presentations/skills/presentations/assets/builtin_templates/grid-layout-library/artifact-tool-compose/index.mjs",
  "skills/presentations/skills/presentations/assets/builtin_templates/grid-layout-library/assets/previews/layout-library.png",
  "skills/office-kit/.codex-plugin/plugin.json",
  "skills/office-kit/README.md",
  "skills/office-kit/skills/office-kit/SKILL.md",
  "skills/office-kit/skills/office-kit/agents/openai.yaml",
  "skills/office-kit/skills/office-kit/references/routing.md",
  "skills/office-kit/skills/office-kit/references/template-selection.md",
  "skills/office-kit/skills/office-kit/references/repl.md",
  "skills/office-kit/skills/office-kit/references/review.md",
  "skills/template-creator/.codex-plugin/plugin.json",
  "skills/template-creator/manifest.json",
  "skills/template-creator/README.md",
  "skills/template-creator/assets/icon.svg",
  "skills/template-creator/skills/template-creator/SKILL.md",
  "skills/template-creator/skills/template-creator/agents/agent.yaml",
  "skills/template-creator/skills/template-creator/assets/icon.svg",
  "skills/template-creator/skills/template-creator/manifest.txt",
  "skills/template-creator/skills/template-creator/scripts/create-template-skill.mjs",
  "skills/default-template-library/LICENSE.md",
  "skills/default-template-library/README.md",
  "skills/default-template-library/manifest.json",
  "skills/default-template-library/integrity.json",
  "skills/default-template-library/skills/artifact-template-business-review/SKILL.md",
  "skills/default-template-library/skills/artifact-template-business-review/artifact-template.json",
  "skills/default-template-library/skills/artifact-template-business-review/assets/reference.pptx",
  "skills/default-template-library/skills/artifact-template-business-review/assets/preview.png",
  "skills/pdf/.codex-plugin/plugin.json",
  "skills/pdf/manifest.json",
  "skills/pdf/README.md",
  "skills/pdf/skills/pdf/SKILL.md",
  "skills/pdf/skills/pdf/agents/openai.yaml",
  "skills/pdf/skills/pdf/agents/agent.yaml",
  "skills/pdf/skills/pdf/manifest.txt",
  "skills/pdf/skills/pdf/artifact_tool/API_QUICK_START.md",
  "skills/pdf/skills/pdf/examples/public-api-end-to-end.mjs",
  "skills/pdf/skills/pdf/examples/accessible-board-report.mjs",
  "skills/pdf/skills/pdf/examples/provider-workflows.md",
  "skills/pdf/skills/pdf/examples/reportlab-report-spec.json",
  "skills/pdf/skills/pdf/examples/pymupdf-edit-operations.json",
  "skills/pdf/skills/pdf/examples/pymupdf-redaction-operations.json",
  "skills/pdf/skills/pdf/examples/pymupdf-ocr-redaction-operations.json",
  "skills/pdf/skills/pdf/references/PROVIDER_MATRIX.md",
  "skills/pdf/skills/pdf/references/SAVE_POLICIES.md",
  "skills/pdf/skills/pdf/references/SECURITY_CHECKLIST.md",
  "skills/pdf/skills/pdf/references/PRODUCT_BOUNDARIES.md",
  "skills/pdf/skills/pdf/references/AUDIT_SCHEMA.md",
  "skills/pdf/skills/pdf/references/pdf-audit-v1.schema.json",
  "skills/pdf/skills/pdf/scripts/pdf_provider.py",
  "skills/pdf/skills/pdf/scripts/qpdf_provider.py",
  "skills/pdf/skills/pdf/scripts/pyhanko_provider.py",
  "skills/pdf/skills/pdf/scripts/pyhanko_sign_provider.py",
  "skills/pdf/skills/pdf/scripts/verapdf_provider.py",
  "skills/pdf/skills/pdf/scripts/ocrmypdf_provider.py",
  "skills/pdf/skills/pdf/scripts/pikepdf_provider.py",
  "skills/pdf/skills/pdf/scripts/mupdf.mjs",
  "skills/pdf/skills/pdf/scripts/reportlab_create.py",
  "skills/pdf/skills/pdf/scripts/pdfplumber_extract.py",
  "skills/pdf/skills/pdf/scripts/pypdf_edit.py",
  "skills/pdf/skills/pdf/scripts/pymupdf_edit.py",
  "skills/pdf/skills/pdf/scripts/residue_scan.py",
  "skills/pdf/skills/pdf/scripts/pdf_audit.py",
  "skills/pdf/skills/pdf/scripts/python_runtime.py",
  "skills/pdf/skills/pdf/tasks/create.md",
  "skills/pdf/skills/pdf/tasks/read_review.md",
  "skills/pdf/skills/pdf/tasks/edit_existing.md",
  "skills/pdf/skills/pdf/tasks/forms_annotations.md",
  "skills/pdf/skills/pdf/tasks/sign_verify.md",
  "skills/pdf/skills/pdf/tasks/redact.md",
  "skills/pdf/skills/pdf/tasks/accessibility.md",
  "skills/pdf/skills/pdf/tasks/render_review.md",
  "skills/pdf/skills/pdf/tasks/provider_setup.md",
  "skills/pdf/skills/pdf/tasks/repair_linearize.md",
  "skills/pdf/skills/pdf/tasks/encryption.md",
  "skills/pdf/skills/pdf/tasks/ocr.md",
  "skills/pdf/skills/pdf/tasks/structure_clean.md",
]) {
  assert.ok(files.includes(required), `npm package is missing ${required}`);
}
assert.equal(
  files.includes("skills/spreadsheets/.app.json"),
  false,
  "npm package must not ship the retired host-connector declaration",
);
assert.ok(files.every((file) => !file.includes("/bin/") && !file.includes("/obj/")), "npm package must exclude dotnet bin/obj build output");
for (const removed of [
  "src/codecs/office-codec-policy.mjs",
  "skills/shared/office-kit-compat.mjs",
  "src/spreadsheet/ooxml-drawings.mjs",
  "src/spreadsheet/ooxml-pivots.mjs",
  "src/presentation/master-graph.mjs",
  "src/presentation/opaque-objects.mjs",
  "src/presentation/ooxml-picture-bullets.mjs",
]) assert.ok(!files.includes(removed), `npm package must not contain removed legacy Office implementation ${removed}`);
assert.ok(!files.includes("skills/reference-sync.json"), "npm package must exclude the repository-only reference source snapshot");
assert.ok(files.every((file) => !file.includes("/tests/") && !file.startsWith("test/")), "npm package must exclude development-only test sources");
assert.ok(files.every((file) => !file.includes(".DS_Store") && !file.includes("__pycache__") && !file.endsWith(".pyc")), "npm package must exclude local metadata and Python bytecode");
assert.ok(files.filter((file) => file.startsWith("src/pdf/providers/")).every((file) => !/\.(?:tar\.gz|tgz|zip|whl|jar|exe|dylib|so)$/i.test(file)), "npm package must ship provider policy/source only, never capability-pack binaries");
assert.ok(files.every((file) => !file.startsWith("reference/")), "npm package must exclude reference material");
assert.ok(!files.includes("native/OfficeBridge/OfficeBridge.sln"), "npm package must not publish a solution whose test project is repository-only");
const packagedTemplateSidecars = files.filter((file) =>
  /^skills\/default-template-library\/skills\/artifact-template-[^/]+\/artifact-template\.json$/u.test(file),
);
assert.equal(
  packagedTemplateSidecars.length,
  20,
  "npm package must ship exactly the 20 audited default templates",
);
assert.ok(files.every((file) => !file.startsWith("native/OfficeKit/") && !file.startsWith("scripts/")), "npm runtime package must not duplicate repository-only OfficeKit source or build tooling");
assert.ok(
  files.every((file) => !file.startsWith("standalone/")),
  "npm runtime package must not contain platform-specific standalone release assets",
);
assert.ok(files.every((file) => !file.startsWith("evals/") && file !== "docs/agent-evals.md"), "npm runtime package must exclude the evaluator-side PromptBench and its oracle documentation");
assert.ok(!files.includes("docs/coverage.md") && !files.includes("docs/release.md") && !files.includes("docs/reference-runtime-architecture.md") && !files.includes("native/OfficeKit/README.md"), "npm runtime package must exclude repository-only coverage, release history, and subsystem implementation notes");
const skillPngs = report.files.filter(({ path: filename }) => /^skills\/(?:documents|spreadsheets|presentations|pdf)\/.*\.png$/.test(filename));
const skillPngBytes = skillPngs.reduce((total, { size }) => total + size, 0);
assert.equal(skillPngs.length, 40, "npm package must retain all 40 public Skill PNG assets");
assert.ok(skillPngBytes < maxSkillPngBytes, `public Skill PNG payload unexpectedly large: ${skillPngBytes} (limit ${maxSkillPngBytes})`);
assert.ok(report.size < maxPackedBytes, `npm package archive unexpectedly large: ${report.size} (limit ${maxPackedBytes})`);
assert.ok(report.unpackedSize < maxUnpackedBytes, `npm package unpacked size unexpectedly large: ${report.unpackedSize} (limit ${maxUnpackedBytes})`);

console.log("package contents smoke ok");
