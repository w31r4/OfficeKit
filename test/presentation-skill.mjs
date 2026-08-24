import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

import { DocumentFile, DocumentModel, FileBlob, Presentation, PresentationFile, SpreadsheetFile, Workbook } from "../src/index.mjs";
import {
  generateOfficeInput,
  PPTX_RICH_NOTES_FIXTURE,
  PPTX_TITLE_NOTES_FIXTURE,
} from "../scripts/agent-eval-office-fixtures.mjs";
import {
  nativePresentationRenderStatus,
  runPresentationFixture,
  verifyPresentationFile,
} from "./skill-harness/presentations/scripts/workflow.mjs";
import { editImportedObjectAccessibility } from "../skills/presentations/skills/presentations/examples/officekit-object-accessibility-edit-workflow.mjs";

const fixtureDir = path.join("test", "skill-harness", "presentations", "fixtures");
const repoRoot = path.resolve(import.meta.dirname, "..");
const presentationSkillDir = path.join(repoRoot, "skills", "presentations", "skills", "presentations");
const packagedSlideRenderer = path.join(presentationSkillDir, "container_tools", "render_slides.py");
const packagedRasterHelper = path.join(presentationSkillDir, "container_tools", "ensure_raster_image.py");
const [packagedSlideRendererSource, packagedRasterHelperSource, presentationSkillEntrySource, advancedImportedSource] = await Promise.all([
  fs.readFile(packagedSlideRenderer, "utf8"),
  fs.readFile(packagedRasterHelper, "utf8"),
  fs.readFile(path.join(presentationSkillDir, "SKILL.md"), "utf8"),
  fs.readFile(path.join(presentationSkillDir, "references", "advanced-imported-editing.md"), "utf8"),
]);
const presentationSkillSource = `${presentationSkillEntrySource}\n${advancedImportedSource}`;
const templateFollowingSource = await fs.readFile(path.join(presentationSkillDir, "references", "template-following.md"), "utf8");
const templateConditionedSource = await fs.readFile(path.join(presentationSkillDir, "references", "template-conditioned-generation.md"), "utf8");
const designMechanismsSource = await fs.readFile(path.join(presentationSkillDir, "references", "design-mechanisms.md"), "utf8");
const audienceTextSource = await fs.readFile(path.join(presentationSkillDir, "references", "audience-text-editing.md"), "utf8");
const designReviewSource = await fs.readFile(path.join(presentationSkillDir, "references", "design-review.md"), "utf8");
const authoringPlanSource = await fs.readFile(path.join(presentationSkillDir, "references", "authoring-plan.md"), "utf8");
const taskNames = ["create", "create-from-template", "edit-existing", "continue", "review-deliver"];
const taskSources = Object.fromEntries(await Promise.all(taskNames.map(async (name) => [
  name,
  await fs.readFile(path.join(presentationSkillDir, "tasks", `${name}.md`), "utf8"),
])));
assert.ok(presentationSkillEntrySource.split(/\r?\n/u).length <= 350, "the Presentation Skill entrypoint must stay within 350 lines");
for (const name of taskNames) assert.match(presentationSkillEntrySource, new RegExp(`tasks/${name.replace(/[.*+?^${}()|[\\]\\]/g, "\\$&")}[.]md`, "u"));
for (const sourceMode of ["self-directed", "design-system", "template", "style-transfer"]) assert.match(presentationSkillEntrySource, new RegExp(`\\b${sourceMode}\\b`, "u"));
for (const mechanism of ["editorial-minimal", "enterprise-data-review", "technical-architecture", "visual-narrative", "academic-research", "brand-launch"]) assert.match(designMechanismsSource, new RegExp(`\\b${mechanism}\\b`, "u"));
assert.match(authoringPlanSource, /ctx\.plan\(revised, \{ expectedSha256: first\.sha256 \}\)/u);
assert.match(audienceTextSource, /Lock facts and sources[\s\S]*Rewrite for the audience[\s\S]*Compress against the real page[\s\S]*Review the deck voice/u);
assert.match(designReviewSource, /bounded signals[\s\S]*intentionalWarnings[\s\S]*concrete[\s\S]*warning-free report as proof of good\s+design/u);
assert.match(taskSources.create, /Shipped route A[\s\S]*Experimental compiler route C[\s\S]*Do not switch routes after a failure/u);
assert.match(taskSources.create, /minimumBodyFontSize[\s\S]*20[\s\S]*minimumCaptionFontSize[\s\S]*18[\s\S]*four different silhouettes/u);
assert.match(taskSources.create, /dominant reading anchor[\s\S]*quantitative\s+claim[\s\S]*clear contrast/u);
assert.match(taskSources.create, /palette\.strict:\s*false[\s\S]*closed brand palette[\s\S]*theme defaults/u);
assert.match(taskSources.create, /no unrecorded[\s\S]*cardWallPattern[\s\S]*intentionalWarnings[\s\S]*concrete[\s\S]*reason/u);
assert.match(taskSources["edit-existing"], /changedPageIds[\s\S]*undeclared page change/u);
assert.doesNotMatch(packagedSlideRendererSource, /pdf2image/i, "the packaged slide renderer must not require an undeclared Python package");
assert.doesNotMatch(packagedRasterHelperSource, /pdf2image/i, "the packaged raster helper must not require an undeclared Python package");
assert.match(presentationSkillSource, /chartDataValue[\s\S]*ChartPart cache[\s\S]*workbook cell/i);
assert.match(presentationSkillSource, /diagramText[\s\S]*token-splicing[\s\S]*a:t/i);
assert.match(presentationSkillSource, /does not authorize chart or diagram identity, relationships,[\s\S]*series or SmartArt topology/i);
assert.match(presentationSkillSource, /svgEditCapability[\s\S]*getSvgEditLeaves[\s\S]*editSvgLeaf/i);
assert.match(presentationSkillSource, /RGB fill\/stroke[\s\S]*opacity in `0\.\.1`[\s\S]*translate[\s\S]*scale[\s\S]*rotate/i);
assert.match(presentationSkillSource, /reinspect before another edit[\s\S]*bound to the current\s+image bytes[\s\S]*imported package revision/i);
assert.match(presentationSkillSource, /Stylesheets[\s\S]*external references[\s\S]*foreignObject[\s\S]*XML selectors[\s\S]*outside\s+this\s+capability/i);
assert.match(presentationSkillSource, /reimport and verify[\s\S]*new-media footprint[\s\S]*byte-identical non-target parts[\s\S]*affected slide render/i);
assert.match(templateFollowingSource, /presentation\.designProfile\(\{ maxItems: 64 \}\)[\s\S]*descriptive evidence[\s\S]*not permission to edit raw XML/i);
assert.match(presentationSkillSource, /getSvgTextNodes\(\)[\s\S]*node\.text[\s\S]*not[\s\S]*node\.value[\s\S]*editSvgText[\s\S]*value/i);
assert.match(presentationSkillSource, /slide\.continuationCapability[\s\S]*pending-clone[\s\S]*export\/reimport[\s\S]*bounded-overlay[\s\S]*textbox[\s\S]*embedded rectangular[\s\S]*clean SlidePart export[\s\S]*Commit\/reopen/i);
assert.match(templateConditionedSource, /continuationCapability[\s\S]*pending-clone[\s\S]*export\/reimport/i);
assert.match(templateConditionedSource, /continuationCapability[\s\S]*bounded-overlay[\s\S]*clean export[\s\S]*commit\/reopen/i);
assert.match(templateFollowingSource, /continuationCapability[\s\S]*export\/reimport[\s\S]*pending-clone[\s\S]*bounded-overlay/i);
assert.match(templateFollowingSource, /separate revision[\s\S]*continuationCapability[\s\S]*arbitrary tables, charts, connectors, groups, or native nodes/i);

const root = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-presentation-skill-test-"));
const baselineDir = path.join(root, "baselines");
const nativeStatus = nativePresentationRenderStatus();
const nativeRender = nativeStatus.available ? "required" : "auto";

function itemByName(items, name) {
  const item = items.find((candidate) => candidate.name === name);
  assert.ok(item, "Missing presentation skill object " + name);
  return item;
}

function relationshipPartPath(partPath) {
  const slash = partPath.lastIndexOf("/");
  const directory = slash < 0 ? "" : partPath.slice(0, slash + 1);
  const fileName = slash < 0 ? partPath : partPath.slice(slash + 1);
  return `${directory}_rels/${fileName}.rels`;
}

function nativeNonVisualProperties(xml, name) {
  const record = [...xml.matchAll(/<p:cNvPr\b[^>]*(?:\/>|>[\s\S]*?<\/p:cNvPr>)/g)]
    .map((match) => match[0])
    .find((candidate) => candidate.includes(`name="${name}"`));
  assert.ok(record, `Missing native p:cNvPr for ${name}`);
  return record;
}

function accessibilityLocatorByName(presentation, name, objectKind) {
  const inspection = presentation.inspect({ kind: "shape,connector,groupShape,image,table,chart", maxChars: 2_000_000 });
  assert.equal(inspection.truncated, false);
  const matches = inspection.ndjson
    .split("\n")
    .filter(Boolean)
    .map((line) => JSON.parse(line))
    .filter((record) => record.name === name && (record.kind === "groupShape" ? "group" : record.kind) === objectKind);
  assert.equal(matches.length, 1, `Expected one ${objectKind} named ${name}`);
  const [record] = matches;
  return {
    slide: record.slide,
    id: record.id,
    objectKind,
    name: record.name,
    ...(record.parentGroupId === undefined ? {} : { parentGroupId: record.parentGroupId }),
  };
}

try {
  const readiness = await runPresentationFixture(path.join(fixtureDir, "agent-readiness.json"), {
    outputDir: path.join(root, "agent-readiness"),
    nativeRender,
    baselineDir,
    writeBaseline: true,
  });
  assert.equal(Object.hasOwn(readiness, "roundtripCodec"), false);
  assert.equal(readiness.qa.verify.ok, true);
  assert.equal(readiness.qa.summary.packageOk, true);
  assert.equal(readiness.qa.presentation.slides.count, 3);
  assert.equal(readiness.qa.modelRender.slides.length, 3);
  assert.equal(readiness.qa.modelRender.montage.ok, true);
  assert.equal(readiness.qa.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  if (nativeStatus.available) assert.equal(readiness.qa.nativeRender.pageCount, 3);

  const renderPython = process.env.OFFICE_KIT_PRESENTATIONS_PYTHON || "python3";
  const pythonVersion = spawnSync(renderPython, ["-S", "--version"], { encoding: "utf8" });
  if (pythonVersion.status === 0) {
    const packagedRenderDir = path.join(root, "packaged-slide-renderer");
    const renderedPptx = spawnSync(renderPython, [
      "-S",
      packagedSlideRenderer,
      readiness.pptxPath,
      "--output_dir",
      path.join(packagedRenderDir, "pptx"),
    ], {
      cwd: repoRoot,
      encoding: "utf8",
      timeout: 60_000,
      env: { ...process.env, PYTHONNOUSERSITE: "1" },
    });
    assert.equal(renderedPptx.status, 0, `The packaged slide renderer must render PPTX without third-party Python imports\nSTDOUT:\n${renderedPptx.stdout}\nSTDERR:\n${renderedPptx.stderr}`);
    assert.deepEqual((await fs.readdir(path.join(packagedRenderDir, "pptx"))).filter((name) => /^slide-\d+\.png$/.test(name)).sort(), ["slide-1.png", "slide-2.png", "slide-3.png"]);

    if (nativeStatus.available) {
      const nativePdfPath = readiness.qa.nativeRender.pdfPath;
      assert.equal(typeof nativePdfPath, "string");
      const renderedPdf = spawnSync(renderPython, [
        "-S",
        packagedSlideRenderer,
        nativePdfPath,
        "--output_dir",
        path.join(packagedRenderDir, "pdf"),
      ], {
        cwd: repoRoot,
        encoding: "utf8",
        timeout: 60_000,
        env: { ...process.env, PYTHONNOUSERSITE: "1" },
      });
      assert.equal(renderedPdf.status, 0, `The packaged slide renderer must use Poppler directly for PDF input\nSTDOUT:\n${renderedPdf.stdout}\nSTDERR:\n${renderedPdf.stderr}`);
      assert.deepEqual((await fs.readdir(path.join(packagedRenderDir, "pdf"))).filter((name) => /^slide-\d+\.png$/.test(name)).sort(), ["slide-1.png", "slide-2.png", "slide-3.png"]);

      const rasterDir = path.join(packagedRenderDir, "raster");
      const rasterized = spawnSync(renderPython, [
        "-S",
        packagedRasterHelper,
        "--input_files",
        nativePdfPath,
        "--output_dir",
        rasterDir,
      ], {
        cwd: repoRoot,
        encoding: "utf8",
        timeout: 60_000,
        env: { ...process.env, PYTHONNOUSERSITE: "1" },
      });
      assert.equal(rasterized.status, 0, `The packaged raster helper must use Poppler directly for PDF input\nSTDOUT:\n${rasterized.stdout}\nSTDERR:\n${rasterized.stderr}`);
      assert.ok((await fs.stat(path.join(rasterDir, "native-render.png"))).size > 0);
    }
  }

  const workflowSlide = readiness.qa.presentation.slides.getItem(0);
  assert.deepEqual(workflowSlide.background, { fill: "#f1f5f9", mode: "solid" });
  assert.deepEqual(readiness.qa.presentation.slides.getItem(1).background, { fill: "#fff7ed", mode: "solid" });
  const authoredCard = itemByName(workflowSlide.shapes.items, "author-card");
  assert.equal(authoredCard.geometry, "roundRect");
  assert.deepEqual(authoredCard.accessibility, {
    title: "Create stage",
    description: "First stage of the workflow, where the presentation is authored.",
    decorative: false,
  });
  assert.deepEqual(authoredCard.accessibilityCapability, { sourceBound: true, editable: true, addable: true });
  const workflowMatrix = itemByName(workflowSlide.tables.items, "workflow-matrix");
  assert.deepEqual(workflowMatrix.accessibility, {
    title: "Workflow verification matrix",
    description: "Table listing semantic and visual review gates with their current states.",
  });
  assert.deepEqual(workflowMatrix.accessibilityCapability, { sourceBound: true, editable: true, addable: true });
  const readinessBar = itemByName(readiness.qa.presentation.slides.getItem(1).charts.items, "readiness-bar");
  assert.deepEqual(readinessBar.accessibility, {
    title: "Readiness scores by gate",
    description: "Bar chart showing Create at 78, Inspect at 92, and Render at 85.",
  });
  assert.deepEqual(readinessBar.accessibilityCapability, { sourceBound: true, editable: true, addable: true });
  const workflowEvidenceImage = itemByName(workflowSlide.images.items, "workflow-evidence-image");
  assert.deepEqual(workflowEvidenceImage.accessibility, {
    title: "Workflow evidence image",
    description: "Status image confirming the OfficeKit workflow evidence.",
  });
  assert.equal(workflowEvidenceImage.alt, workflowEvidenceImage.accessibility.description);
  assert.deepEqual(workflowEvidenceImage.accessibilityCapability, { sourceBound: true, editable: true, addable: true });

  const accessibilityEditDir = path.join(root, "object-accessibility-edits");
  await fs.mkdir(accessibilityEditDir, { recursive: true });
  const accessibilityEditCases = [
    { key: "shape", locator: accessibilityLocatorByName(readiness.qa.presentation, "author-card", "shape"), update: { description: "Reviewed create-stage description." } },
    { key: "connector", locator: accessibilityLocatorByName(readiness.qa.presentation, "create-to-verify", "connector"), update: { title: "Reviewed create-to-verify flow" } },
    { key: "group", locator: accessibilityLocatorByName(readiness.qa.presentation, "nested-qa-group", "group"), update: { title: "Reviewed nested render verification" } },
    { key: "image", locator: accessibilityLocatorByName(readiness.qa.presentation, "workflow-evidence-image", "image"), update: { description: "Reviewed status image for the OfficeKit workflow." } },
    { key: "table", locator: accessibilityLocatorByName(readiness.qa.presentation, "workflow-matrix", "table"), update: { title: "Reviewed workflow verification matrix" } },
    { key: "chart", locator: accessibilityLocatorByName(readiness.qa.presentation, "readiness-bar", "chart"), update: { description: "Reviewed readiness scores for Create, Inspect, and Render." } },
  ];
  const readinessSourceBytes = await fs.readFile(readiness.pptxPath);
  for (const testCase of accessibilityEditCases) {
    const sourceTarget = readiness.qa.presentation.resolve(testCase.locator.id);
    assert.ok(sourceTarget);
    const outputPath = path.join(accessibilityEditDir, `${testCase.key}.pptx`);
    const auditPath = path.join(accessibilityEditDir, `${testCase.key}.json`);
    const result = await editImportedObjectAccessibility({
      inputPath: readiness.pptxPath,
      outputPath,
      auditPath,
      locator: testCase.locator,
      expectedAccessibility: structuredClone(sourceTarget.accessibility || {}),
      update: testCase.update,
    });
    assert.equal(result.audit.schema, "office-kit.pptx-object-accessibility-edit-audit.v1");
    assert.deepEqual(result.audit.operation.locator, testCase.locator);
    assert.deepEqual(result.audit.validation.package.changedPaths, [`ppt/slides/slide${testCase.locator.slide}.xml`]);
    assert.equal(result.audit.validation.package.exactlyTargetSlidePartChanged, true);
    assert.equal(result.audit.validation.package.allOtherPartsByteIdentical, true);
    assert.equal(result.audit.validation.model.nonTargetPresentationStatePreserved, true);
    assert.equal(result.audit.validation.modelRender.visuallyUnchanged, true);
    assert.equal(result.audit.validation.accessibilityAudit.targetMachineIssueAbsent, true);
    assert.equal(result.audit.validation.accessibilityAudit.conformanceClaimed, false);
    assert.deepEqual(JSON.parse(await fs.readFile(auditPath, "utf8")), result.audit);
    const roundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(outputPath), {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      name: `${testCase.key}.pptx`,
    }));
    assert.deepEqual(roundTrip.resolve(testCase.locator.id).accessibility, result.audit.operation.result);
    if (testCase.key === "image") assert.equal(roundTrip.resolve(testCase.locator.id).alt, result.audit.operation.result.description);
    assert.deepEqual(await fs.readFile(readiness.pptxPath), readinessSourceBytes);
  }

  const shapeEditCase = accessibilityEditCases[0];
  const shapeExpectedAccessibility = structuredClone(readiness.qa.presentation.resolve(shapeEditCase.locator.id).accessibility);
  await assert.rejects(
    () => editImportedObjectAccessibility({
      inputPath: readiness.pptxPath,
      outputPath: path.join(accessibilityEditDir, "stale.pptx"),
      auditPath: path.join(accessibilityEditDir, "stale.json"),
      locator: shapeEditCase.locator,
      expectedAccessibility: { title: "Stale title" },
      update: { title: "Replacement" },
    }),
    /does not match the expected complete source state/,
  );
  const groupEditCase = accessibilityEditCases[2];
  await assert.rejects(
    () => editImportedObjectAccessibility({
      inputPath: readiness.pptxPath,
      outputPath: path.join(accessibilityEditDir, "wrong-parent.pptx"),
      auditPath: path.join(accessibilityEditDir, "wrong-parent.json"),
      locator: { ...groupEditCase.locator, parentGroupId: "not-the-source-group" },
      expectedAccessibility: structuredClone(readiness.qa.presentation.resolve(groupEditCase.locator.id).accessibility),
      update: { title: "Replacement" },
    }),
    /does not match the complete audit locator/,
  );
  await assert.rejects(
    () => editImportedObjectAccessibility({
      inputPath: readiness.pptxPath,
      outputPath: path.join(accessibilityEditDir, "no-op.pptx"),
      auditPath: path.join(accessibilityEditDir, "no-op.json"),
      locator: shapeEditCase.locator,
      expectedAccessibility: shapeExpectedAccessibility,
      update: { title: shapeExpectedAccessibility.title },
    }),
    /does not change the selected object accessibility state/,
  );
  const connectorEditCase = accessibilityEditCases[1];
  await assert.rejects(
    () => editImportedObjectAccessibility({
      inputPath: readiness.pptxPath,
      outputPath: path.join(accessibilityEditDir, "unclassified.pptx"),
      auditPath: path.join(accessibilityEditDir, "unclassified.json"),
      locator: connectorEditCase.locator,
      expectedAccessibility: structuredClone(readiness.qa.presentation.resolve(connectorEditCase.locator.id).accessibility),
      update: { title: null, description: null, decorative: null },
    }),
    /must leave the selected object explicitly decorative or with an accessibility title\/description/,
  );
  const irregularAccessibilityPath = path.join(accessibilityEditDir, "irregular-source.pptx");
  const irregularAccessibilityZip = await JSZip.loadAsync(readinessSourceBytes);
  const irregularAccessibilityXml = (await irregularAccessibilityZip.file("ppt/slides/slide1.xml").async("text"))
    .replace(/(<p:cNvPr\b[^>]*\bname="author-card")/, '$1 xmlns:fixture="urn:office-kit:skill-accessibility" fixture:opaque="kept"');
  irregularAccessibilityZip.file("ppt/slides/slide1.xml", irregularAccessibilityXml);
  await fs.writeFile(irregularAccessibilityPath, await irregularAccessibilityZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  await assert.rejects(
    () => editImportedObjectAccessibility({
      inputPath: irregularAccessibilityPath,
      outputPath: path.join(accessibilityEditDir, "irregular-output.pptx"),
      auditPath: path.join(accessibilityEditDir, "irregular-audit.json"),
      locator: shapeEditCase.locator,
      expectedAccessibility: {},
      update: { title: "Do not flatten opaque metadata" },
    }),
    /not an editable source-bound p:cNvPr profile/,
  );
  const collisionPath = path.join(accessibilityEditDir, "collision.pptx");
  await fs.writeFile(collisionPath, "occupied");
  await assert.rejects(
    () => editImportedObjectAccessibility({
      inputPath: readiness.pptxPath,
      outputPath: collisionPath,
      auditPath: path.join(accessibilityEditDir, "collision.json"),
      locator: shapeEditCase.locator,
      expectedAccessibility: shapeExpectedAccessibility,
      update: { title: "Replacement" },
    }),
    /outputPath already exists/,
  );
  assert.equal(await fs.access(path.join(accessibilityEditDir, "stale.pptx")).then(() => true, () => false), false);
  assert.equal(await fs.access(path.join(accessibilityEditDir, "wrong-parent.pptx")).then(() => true, () => false), false);
  assert.equal(await fs.access(path.join(accessibilityEditDir, "no-op.pptx")).then(() => true, () => false), false);
  assert.equal(await fs.access(path.join(accessibilityEditDir, "unclassified.pptx")).then(() => true, () => false), false);
  assert.equal(await fs.access(path.join(accessibilityEditDir, "irregular-output.pptx")).then(() => true, () => false), false);

  const cliOutputPath = path.join(accessibilityEditDir, "cli-shape.pptx");
  const cliAuditPath = path.join(accessibilityEditDir, "cli-shape.json");
  const accessibilityCli = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/examples/officekit-object-accessibility-edit-workflow.mjs",
    readiness.pptxPath,
    cliOutputPath,
    cliAuditPath,
    JSON.stringify(shapeEditCase.locator),
    JSON.stringify(shapeExpectedAccessibility),
    JSON.stringify({ title: "CLI-reviewed create stage" }),
  ], { encoding: "utf8" });
  assert.equal(accessibilityCli.status, 0, `object accessibility CLI failed\n${accessibilityCli.stdout}\n${accessibilityCli.stderr}`);
  assert.deepEqual(JSON.parse(accessibilityCli.stdout).changedPaths, ["ppt/slides/slide1.xml"]);
  assert.equal((await fs.stat(cliOutputPath)).isFile(), true);
  assert.equal((await fs.stat(cliAuditPath)).isFile(), true);

  assert.deepEqual(authoredCard.shadow, { color: "#000000", blurRadius: 10, distance: 5, direction: 45, opacity: 0.2 });
  const readinessXml = await (await JSZip.loadAsync(await fs.readFile(readiness.pptxPath))).file("ppt/slides/slide1.xml").async("text");
  const authoredCardNative = nativeNonVisualProperties(readinessXml, "author-card");
  assert.match(authoredCardNative, /\btitle="Create stage"/);
  assert.match(authoredCardNative, /\bdescr="First stage of the workflow, where the presentation is authored\."/);
  assert.match(authoredCardNative, /\{C183D7F6-B498-43B3-948B-1728B52AA6E4\}/);
  assert.match(authoredCardNative, /<adec:decorative\b[^>]*\bval="0"/);
  assert.match(readinessXml, /<p:cNvPr\b(?=[^>]*\bname="workflow-matrix")(?=[^>]*\btitle="Workflow verification matrix")(?=[^>]*\bdescr="Table listing semantic and visual review gates with their current states\.")[^>]*\/>/);
  assert.match(readinessXml, /<p:cNvPr\b(?=[^>]*\bname="workflow-evidence-image")(?=[^>]*\btitle="Workflow evidence image")(?=[^>]*\bdescr="Status image confirming the OfficeKit workflow evidence\.")[^>]*\/>/);
  const readinessChartXml = await (await JSZip.loadAsync(await fs.readFile(readiness.pptxPath))).file("ppt/slides/slide2.xml").async("text");
  assert.match(readinessChartXml, /<p:cNvPr\b(?=[^>]*\bname="readiness-bar")(?=[^>]*\btitle="Readiness scores by gate")(?=[^>]*\bdescr="Bar chart showing Create at 78, Inspect at 92, and Render at 85\.")[^>]*\/>/);
  const verifyCard = itemByName(workflowSlide.shapes.items, "verify-card");
  assert.equal(verifyCard.geometry, "custom");
  assert.equal(verifyCard.customConnectionSites.length, 4);
  assert.deepEqual(verifyCard.customConnectionSites[0], { angle: "3cd4", x: "hc", y: "t" });
  assert.deepEqual(verifyCard.customConnectionSites[3], { angle: 0, x: "r", y: "vc" });
  assert.equal(verifyCard.inspectRecord().customConnectionSiteCount, 4);
  assert.equal(verifyCard.customAdjustmentHandles.length, 2);
  assert.deepEqual(verifyCard.customAdjustmentHandles[0], { kind: "xy", xAdjustment: "adjX", minX: "l", maxX: "r", yAdjustment: "adjY", minY: "t", maxY: "b", x: "hc", y: "vc" });
  assert.equal(verifyCard.inspectRecord().customAdjustmentHandleCount, 2);
  assert.deepEqual(verifyCard.textRectangle, { left: "l", top: "t", right: "r", bottom: "b" });
  assert.equal(Object.hasOwn(verifyCard.customPaths[0], "width"), false);
  assert.equal(Object.hasOwn(verifyCard.customPaths[0], "height"), false);
  assert.deepEqual(verifyCard.customPaths[0].commands[0], { moveTo: { x: "l", y: "t" } });
  assert.equal(itemByName(workflowSlide.shapes.items, "workflow-title").geometry, "textbox");
  const workflowDivider = itemByName(workflowSlide.shapes.items, "free-workflow-divider");
  assert.equal(workflowDivider.geometry, "line");
  assert.deepEqual(workflowDivider.position, { left: 72, top: 122, width: 1108, height: 0 });
  assert.equal(workflowDivider.line.style, "dash-dot");
  assert.deepEqual(workflowDivider.line.head, { type: "oval", width: "sm", length: "med" });
  assert.deepEqual(workflowDivider.line.tail, { type: "arrow", width: "lg", length: "sm" });
  assert.equal(workflowDivider.line.cap, "round");
  assert.equal(workflowDivider.line.join, "bevel");
  assert.match(workflowDivider.toSvg(), /<line\b[^>]*stroke-dasharray="8 4 2 4"/);
  assert.match(workflowDivider.toSvg(), /marker-start=/);
  assert.match(workflowDivider.toSvg(), /marker-end=/);
  assert.equal(itemByName(workflowSlide.tables.items, "workflow-matrix").values[1][1], "Pass");
  const curved = itemByName(workflowSlide.connectors.items, "create-to-verify");
  assert.equal(curved.connectorType, "curved");
  assert.equal(curved.startSiteIndex, 3);
  assert.equal(curved.endSiteIndex, 1);
  assert.deepEqual(curved.head, { type: "arrow", width: "lg", length: "sm" });
  assert.deepEqual(curved.tail, { type: "diamond", width: "sm", length: "lg" });
  assert.equal(curved.line.style, "dashed");
  assert.deepEqual(curved.accessibility, {
    title: "Create to verify flow",
    description: "Connector showing that authored content advances to semantic verification.",
  });
  assert.deepEqual(curved.accessibilityCapability, { sourceBound: true, editable: true, addable: true });
  const elbow = itemByName(workflowSlide.connectors.items, "verify-to-deliver");
  assert.equal(elbow.connectorType, "elbow");
  assert.equal(elbow.line.startArrow, "triangle");
  assert.equal(elbow.line.endArrow, "triangle");
  assert.ok(elbow.startTargetId && elbow.endTargetId);
  assert.deepEqual(elbow.accessibility, { decorative: true });
  assert.deepEqual(elbow.accessibilityCapability, { sourceBound: true, editable: true, addable: true });
  const decorativeConnectorNative = nativeNonVisualProperties(readinessXml, "verify-to-deliver");
  assert.doesNotMatch(decorativeConnectorNative, /\b(?:title|descr)=/);
  assert.match(decorativeConnectorNative, /\{C183D7F6-B498-43B3-948B-1728B52AA6E4\}/);
  assert.match(decorativeConnectorNative, /<adec:decorative\b[^>]*\bval="1"/);
  const charts = readiness.qa.presentation.slides.getItem(1).charts.items;
  assert.deepEqual(charts.map((chart) => chart.chartType), ["bar", "line", "pie"]);
  assert.equal(charts[0].dataLabels.showValue, true);
  assert.equal(charts[1].series[0].marker.symbol, "circle");
  assert.equal(charts[2].dataLabels.showCategoryName, true);
  const groupedWorkflowSlide = readiness.qa.presentation.slides.getItem(2);
  const nativeGroup = itemByName(groupedWorkflowSlide.groups.items, "native-agent-group");
  assert.deepEqual(nativeGroup.childFrame, { left: -80, top: 40, width: 1280, height: 540 });
  assert.deepEqual(nativeGroup.accessibility, {
    title: "Native grouped workflow",
    description: "Group containing the model, OfficeKit format, and render verification stages.",
  });
  assert.deepEqual(nativeGroup.accessibilityCapability, { sourceBound: true, editable: true, addable: true });
  assert.deepEqual(nativeGroup.children.map((child) => child.layoutJson().kind), ["connector", "textbox", "textbox", "textbox", "groupShape"]);
  const nestedQaGroup = itemByName(nativeGroup.groups.items, "nested-qa-group");
  assert.equal(nestedQaGroup.shapes.items[0].text.value, "Render + verify");
  assert.deepEqual(nestedQaGroup.accessibility, {
    title: "Nested render verification",
    description: "Nested group for render and verification evidence.",
  });
  const groupedFlow = itemByName(nativeGroup.connectors.items, "grouped-flow");
  assert.equal(groupedFlow.line.endArrow, "triangle");
  assert.deepEqual(groupedFlow.accessibility, {
    title: "Model to OfficeKit flow",
    description: "Connector from the JavaScript model stage to the OfficeKit format stage.",
  });
  assert.match(readiness.qa.inspect.ndjson, /OfficeKit closes the presentation loop/);
  assert.match(readiness.qa.inspect.ndjson, /free-workflow-divider/);
  assert.match(readiness.qa.inspect.ndjson, /Coverage mix/);
  assert.match(readiness.qa.inspect.ndjson, /native-agent-group/);

  const readinessZip = await JSZip.loadAsync(await fs.readFile(readiness.pptxPath));
  const firstSlideXml = await readinessZip.file("ppt/slides/slide1.xml").async("text");
  assert.match(firstSlideXml, /<a:srgbClr val="F1F5F9"/);
  assert.match(firstSlideXml, /<a:prstGeom prst="roundRect"[^>]*>/);
  assert.match(firstSlideXml, /<a:custGeom\b/);
  assert.match(firstSlideXml, /<a:ahLst><a:ahXY gdRefX="adjX" minX="l" maxX="r" gdRefY="adjY" minY="t" maxY="b"><a:pos x="hc" y="vc"\s*\/><\/a:ahXY><a:ahPolar gdRefR="adjRadius" minR="l" maxR="ss" gdRefAng="adjAngle" minAng="t" maxAng="cd2"><a:pos x="r" y="vc"\s*\/><\/a:ahPolar><\/a:ahLst>/);
  assert.match(firstSlideXml, /<a:cxnLst><a:cxn ang="3cd4"><a:pos x="hc" y="t"\s*\/>/);
  assert.match(firstSlideXml, /<a:cxn ang="0"><a:pos x="r" y="vc"\s*\/><\/a:cxn><\/a:cxnLst>/);
  assert.match(firstSlideXml, /<a:rect l="l" t="t" r="r" b="b"\s*\/>/);
  assert.match(firstSlideXml, /<a:path><a:moveTo><a:pt x="l" y="t"\s*\/>/);
  assert.match(firstSlideXml, /<a:prstGeom\b[^>]*\bprst="line"/);
  assert.match(firstSlideXml, /<a:ext\b[^>]*\bcx="10553700"[^>]*\bcy="0"/);
  assert.match(firstSlideXml, /<a:ln\b[^>]*cap="rnd"/);
  assert.match(firstSlideXml, /<a:headEnd\b[^>]*type="oval"[^>]*w="sm"[^>]*len="med"/);
  assert.match(firstSlideXml, /<a:tailEnd\b[^>]*type="arrow"[^>]*w="lg"[^>]*len="sm"/);
  assert.match(firstSlideXml, /<p:cNvSpPr txBox="1"\s*\/>/);
  assert.match(firstSlideXml, /<p:cxnSp>/);
  assert.match(firstSlideXml, /<p:cNvPr\b(?=[^>]*\bname="create-to-verify")(?=[^>]*\btitle="Create to verify flow")(?=[^>]*\bdescr="Connector showing that authored content advances to semantic verification\.")[^>]*\/>/);
  assert.match(firstSlideXml, /<a:stCxn\b[^>]*idx="3"/);
  assert.match(firstSlideXml, /<a:endCxn\b[^>]*idx="1"/);
  assert.match(firstSlideXml, /prst="curvedConnector3"/);
  assert.match(firstSlideXml, /<a:prstDash val="dash"/);
  const groupedSlideXml = await readinessZip.file("ppt/slides/slide3.xml").async("text");
  assert.equal((groupedSlideXml.match(/<p:grpSp>/g) || []).length, 2);
  assert.match(groupedSlideXml, /<a:chOff x="-762000" y="381000"\s*\/>/);
  assert.match(groupedSlideXml, /<p:cNvPr\b(?=[^>]*\bname="native-agent-group")(?=[^>]*\btitle="Native grouped workflow")(?=[^>]*\bdescr="Group containing the model, OfficeKit format, and render verification stages\.")[^>]*\/>/);
  assert.match(groupedSlideXml, /<p:cNvPr\b(?=[^>]*\bname="nested-qa-group")(?=[^>]*\btitle="Nested render verification")(?=[^>]*\bdescr="Nested group for render and verification evidence\.")[^>]*\/>/);
  assert.match(groupedSlideXml, /<p:cNvPr\b(?=[^>]*\bname="grouped-flow")(?=[^>]*\btitle="Model to OfficeKit flow")(?=[^>]*\bdescr="Connector from the JavaScript model stage to the OfficeKit format stage\.")[^>]*\/>/);
  assert.equal(Object.keys(readinessZip.files).filter((name) => /^ppt\/(?:slides\/)?charts\/chart\d+\.xml$/.test(name)).length, 3);

  const compared = await verifyPresentationFile(readiness.pptxPath, {
    outputDir: path.join(root, "agent-readiness-compare"),
    nativeRender,
    baselineDir,
  });
  assert.equal(compared.verify.ok, true);
  assert.equal(compared.modelRender.baselinePageCount, 3);
  assert.equal(compared.modelRender.pageCountMatches, true);
  assert.ok(compared.modelRender.slides.every((slide) => slide.baselineCompared && slide.pixelDiff?.changed === false && slide.ok));
  if (nativeStatus.available) {
    assert.equal(compared.nativeRender.baselinePageCount, 3);
    assert.equal(compared.nativeRender.pageCountMatches, true);
    assert.ok(compared.nativeRender.pages.every((slide) => slide.baselineCompared && slide.pixelDiff?.changed === false && slide.ok));
  }

  const accessibilityAuditDir = path.join(root, "accessibility-audit");
  const accessibilityReportPath = path.join(accessibilityAuditDir, "report.json");
  const readinessSourceBeforeAudit = await fs.readFile(readiness.pptxPath);
  const { auditPptxAccessibility } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-accessibility-audit-workflow.mjs"
  );
  const accessibilityResult = await auditPptxAccessibility({
    inputPath: readiness.pptxPath,
    reportPath: accessibilityReportPath,
    maxChars: 100_000,
  });
  assert.equal(accessibilityResult.report.schema, "office-kit.pptx-accessibility-audit.v1");
  assert.equal(accessibilityResult.report.provider.requested, "office-kit");
  assert.equal(accessibilityResult.report.provider.actual, "office-kit");
  assert.equal(accessibilityResult.report.provider.silentFallback, false);
  assert.deepEqual(accessibilityResult.report.savePolicy, {
    strategy: "none",
    sourceMutation: false,
    artifactProduced: false,
  });
  assert.equal(accessibilityResult.report.operation.type, "presentation-accessibility-audit");
  assert.equal(accessibilityResult.report.accessibility.conformanceClaimed, false);
  assert.equal(accessibilityResult.report.accessibility.manualReviewRequired, true);
  assert.equal(accessibilityResult.report.accessibility.summary.slides, 3);
  assert.equal(accessibilityResult.report.validation.sourceUnchanged, true);
  assert.equal(accessibilityResult.report.validation.presentationVerify.ok, true);
  assert.equal(accessibilityResult.report.boundaries.readingOrder, "manual-native-host-review");
  assert.deepEqual(JSON.parse(await fs.readFile(accessibilityReportPath, "utf8")), accessibilityResult.report);
  assert.deepEqual(await fs.readFile(readiness.pptxPath), readinessSourceBeforeAudit);
  assert.deepEqual(await fs.readdir(accessibilityAuditDir), ["report.json"]);
  await assert.rejects(
    () => auditPptxAccessibility({ inputPath: readiness.pptxPath, reportPath: accessibilityReportPath }),
    /already exists; refusing to overwrite/,
  );
  await assert.rejects(
    () => auditPptxAccessibility({
      inputPath: readiness.pptxPath,
      reportPath: path.join(accessibilityAuditDir, "invalid-limit.json"),
      maxChars: 0,
    }),
    /maxChars must be an integer from 1 through 2000000/,
  );
  assert.equal(await fs.access(path.join(accessibilityAuditDir, "invalid-limit.json")).then(() => true, () => false), false);

  const chartFamiliesDir = path.join(root, "chart-families-workflow");
  const chartFamiliesOutput = path.join(chartFamiliesDir, "chart-families.pptx");
  const chartFamiliesPreview = path.join(chartFamiliesDir, "chart-families.png");
  const chartFamiliesAudit = path.join(chartFamiliesDir, "audit.json");
  const { createAndEditChartFamilyDeck } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-chart-families-workflow.mjs"
  );
  const chartFamiliesResult = await createAndEditChartFamilyDeck({
    outputPath: chartFamiliesOutput,
    previewPath: chartFamiliesPreview,
    auditPath: chartFamiliesAudit,
  });
  assert.equal(chartFamiliesResult.audit.provider.actual, "office-kit");
  assert.equal(chartFamiliesResult.audit.provider.silentFallback, false);
  assert.equal(chartFamiliesResult.audit.operation.type, "greenfield-native-chart-family-author-edit");
  assert.deepEqual(chartFamiliesResult.audit.operation.chartTypes, ["area", "doughnut", "scatter", "bubble"]);
  assert.equal(chartFamiliesResult.audit.validation.verify.ok, true);
  assert.equal(chartFamiliesResult.audit.validation.package.charts.length, 4);
  assert.equal(chartFamiliesResult.audit.validation.package.charts[1].showPercent, true);
  assert.equal(chartFamiliesResult.audit.validation.package.charts[2].numericX, true);
  assert.equal(chartFamiliesResult.audit.validation.package.charts[3].bubbleSizes, true);
  assert.equal((await fs.readFile(chartFamiliesPreview)).subarray(0, 8).toString("hex"), "89504e470d0a1a0a");
  const chartFamiliesRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(chartFamiliesOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "chart-families.pptx",
  }));
  assert.deepEqual(chartFamiliesRoundTrip.slides.getItem(0).charts.items.map((chart) => chart.chartType), ["area", "doughnut", "scatter", "bubble"]);
  assert.equal(itemByName(chartFamiliesRoundTrip.slides.getItem(0).charts.items, "area-family").series[0].values[1], 57);
  assert.equal(itemByName(chartFamiliesRoundTrip.slides.getItem(0).charts.items, "doughnut-family").series[0].values[1], 33);
  assert.equal(itemByName(chartFamiliesRoundTrip.slides.getItem(0).charts.items, "scatter-family").series[0].xValues[1], 22);
  assert.equal(itemByName(chartFamiliesRoundTrip.slides.getItem(0).charts.items, "bubble-family").series[0].bubbleSizes[1], 12);
  const chartFamiliesQa = await verifyPresentationFile(chartFamiliesOutput, {
    outputDir: path.join(chartFamiliesDir, "qa"),
    nativeRender,
  });
  assert.equal(chartFamiliesQa.verify.ok, true);
  assert.equal(chartFamiliesQa.modelRender.slides.length, 1);
  assert.equal(chartFamiliesQa.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  if (nativeStatus.available) assert.equal(chartFamiliesQa.nativeRender.pageCount, 1);

  const chartTrendlineDir = path.join(root, "chart-trendline-workflow");
  const chartTrendlineOutput = path.join(chartTrendlineDir, "chart-trendlines.pptx");
  const chartTrendlinePreview = path.join(chartTrendlineDir, "chart-trendlines.png");
  const chartTrendlineAudit = path.join(chartTrendlineDir, "audit.json");
  const { createAndEditTrendlineDeck } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-chart-trendline-workflow.mjs"
  );
  const chartTrendlineResult = await createAndEditTrendlineDeck({
    outputPath: chartTrendlineOutput,
    previewPath: chartTrendlinePreview,
    auditPath: chartTrendlineAudit,
  });
  assert.equal(chartTrendlineResult.audit.provider.actual, "office-kit");
  assert.equal(chartTrendlineResult.audit.provider.silentFallback, false);
  assert.equal(chartTrendlineResult.audit.operation.type, "greenfield-native-chart-trendline-author-edit");
  assert.deepEqual(chartTrendlineResult.audit.operation.trendlineTypes, ["linear", "movingAvg", "poly", "exp"]);
  assert.deepEqual(chartTrendlineResult.audit.operation.errorBarValueTypes, ["stdDev", "cust"]);
  assert.equal(chartTrendlineResult.audit.validation.verify.ok, true);
  assert.equal((await fs.readFile(chartTrendlinePreview)).subarray(0, 8).toString("hex"), "89504e470d0a1a0a");
  const chartTrendlineRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(chartTrendlineOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "chart-trendlines.pptx",
  }));
  const pipelineTrend = itemByName(chartTrendlineRoundTrip.slides.getItem(0).charts.items, "pipeline-trend");
  const marginTrend = itemByName(chartTrendlineRoundTrip.slides.getItem(0).charts.items, "revenue-margin-trend");
  assert.deepEqual(pipelineTrend.series[0].trendlines.map((trendline) => trendline.type), ["linear", "movingAvg", "poly"]);
  assert.equal(pipelineTrend.series[0].trendlines[0].name, "Updated pipeline projection");
  assert.equal(pipelineTrend.series[0].trendlines[0].forward, 1.5);
  assert.equal(pipelineTrend.series[0].errorBars.value, 2);
  assert.equal(pipelineTrend.series[0].errorBars.line.fill, "#BE123C");
  assert.equal(marginTrend.series[1].trendlines[0].name, "Updated margin projection");
  assert.deepEqual(marginTrend.series[1].errorBars.minusValues, [1, 3, 1, 2]);
  const chartTrendlineZip = await JSZip.loadAsync(await fs.readFile(chartTrendlineOutput));
  const chartTrendlineXml = await Promise.all(Object.keys(chartTrendlineZip.files)
    .filter((name) => /\/charts\/chart\d+\.xml$/.test(name))
    .map((name) => chartTrendlineZip.file(name).async("text")));
  assert.deepEqual(chartTrendlineXml.map((xml) => [...xml.matchAll(/<c:trendlineType val="([^"]+)"\s*\/>/g)].map((match) => match[1])), [["linear", "movingAvg", "poly"], ["exp"]]);
  assert.deepEqual(chartTrendlineXml.map((xml) => [...xml.matchAll(/<c:errValType val="([^"]+)"\s*\/>/g)].map((match) => match[1])), [["stdDev"], ["cust"]]);
  const chartTrendlineQa = await verifyPresentationFile(chartTrendlineOutput, {
    outputDir: path.join(chartTrendlineDir, "qa"),
    nativeRender,
  });
  assert.equal(chartTrendlineQa.verify.ok, true);
  assert.equal(chartTrendlineQa.modelRender.slides.length, 1);
  assert.equal(chartTrendlineQa.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  if (nativeStatus.available) assert.equal(chartTrendlineQa.nativeRender.pageCount, 1);

  const roundtrip = await runPresentationFixture(path.join(fixtureDir, "office-kit-preservation.json"), {
    outputDir: path.join(root, "roundtrip"),
    nativeRender,
  });
  assert.equal(roundtrip.qa.verify.ok, true);
  assert.equal(roundtrip.qa.summary.packageOk, true);
  const roundtripSlide = roundtrip.qa.presentation.slides.getItem(0);
  const editedCard = itemByName(roundtripSlide.shapes.items, "source-card");
  assert.equal(editedCard.text.value, "After edit");
  assert.equal(editedCard.shadow.opacity, 0.35);
  assert.equal(itemByName(roundtripSlide.connectors.items, "editable-connector").line.endArrow, undefined);
  const roundtripTable = itemByName(roundtripSlide.tables.items, "roundtrip-table");
  assert.equal(roundtripTable.values[1][1], "After");
  assert.deepEqual(roundtripTable.mergeRanges, [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }]);
  assert.equal(roundtripTable.getCell(0, 0).columnSpan, 2);
  assert.equal(roundtripTable.getCell(0, 1).editable, false);
  const editedImage = itemByName(roundtripSlide.images.items, "edited-roundtrip-image");
  assert.equal(editedImage.alt, "Roundtrip status after edit");
  assert.deepEqual(editedImage.position, { left: 430, top: 320, width: 170, height: 170 });
  assert.equal(editedImage.fit, "stretch");
  assert.deepEqual(editedImage.crop, { left: 0, top: -0.5, right: 0, bottom: -0.5 });
  const roundtripZip = await JSZip.loadAsync(await fs.readFile(roundtrip.pptxPath));
  const roundtripSlideXml = await roundtripZip.file("ppt/slides/slide1.xml").async("text");
  assert.match(roundtripSlideXml, /<a:srcRect[^>]*t="-50000"/);
  assert.match(roundtripSlideXml, /<a:srcRect[^>]*b="-50000"/);
  assert.match(roundtripSlideXml, /<a:tc gridSpan="2">/);
  assert.match(roundtripSlideXml, /<a:tc hMerge="1">/);
  const editedChart = itemByName(roundtripSlide.charts.items, "roundtrip-chart");
  assert.equal(editedChart.title, "After roundtrip");
  assert.deepEqual(editedChart.series[0].values, [8, 14, 6]);

  const packageDrawing = await runPresentationFixture(path.join(fixtureDir, "package-drawing.json"), {
    outputDir: path.join(root, "package-drawing"),
    nativeRender,
  });
  assert.equal(packageDrawing.qa.verify.ok, true);
  assert.equal(packageDrawing.qa.summary.packageOk, true);
  assert.ok(packageDrawing.qa.packageInspect.parts.some((part) => part.path === "ppt/review/media/agent-status.png"));
  assert.ok(packageDrawing.qa.packageInspect.parts.some((part) => part.path === "ppt/review/charts/agent-readiness.xml"));
  const drawingSlide = packageDrawing.qa.presentation.slides.getItem(0);
  assert.equal(itemByName(drawingSlide.images.items, "Agent status").alt, "Green package status");
  assert.equal(itemByName(drawingSlide.charts.items, "Agent readiness chart").series[0].values[1], 100);

  const coreReview = await runPresentationFixture(path.join(fixtureDir, "modern-comments.json"), {
    outputDir: path.join(root, "core-review"),
    nativeRender,
  });
  assert.equal(coreReview.qa.verify.ok, true);
  assert.equal(coreReview.qa.presentation.slides.getItem(0).comments.items.length, 0);
  assert.equal(itemByName(coreReview.qa.presentation.slides.getItem(0).images.items, "review-status").alt, "Green review status");
  assert.equal(itemByName(coreReview.qa.presentation.slides.getItem(0).connectors.items, "visual-to-delivery").line.endArrow, "triangle");

  const legacyComments = await runPresentationFixture(path.join(fixtureDir, "office-kit-legacy-comments.json"), {
    outputDir: path.join(root, "legacy-comments"),
    nativeRender,
  });
  assert.equal(legacyComments.qa.verify.ok, true);
  assert.equal(legacyComments.qa.summary.packageOk, true);
  const legacyComment = legacyComments.qa.presentation.slides.getItem(0).comments.items[0];
  assert.ok(legacyComment, "legacy comment fixture must survive canonical import/export");
  assert.equal(legacyComment.nativeFormat, "legacy");
  assert.equal(legacyComment.targetId, undefined);
  assert.equal(legacyComment.author, "Presentation Reviewer");
  assert.equal(legacyComment.comments.length, 1);
  assert.equal(legacyComment.comments[0].text, "Confirm the source before delivery.");
  assert.deepEqual(legacyComment.position, { x: 1040, y: 84, unit: "px" });
  assert.match(legacyComments.qa.inspect.ndjson, /Confirm the source before delivery\./);
  const legacyZip = await JSZip.loadAsync(await fs.readFile(legacyComments.pptxPath));
  const commentAuthorsXml = await legacyZip.file("ppt/commentAuthors.xml").async("text");
  const commentXml = await legacyZip.file("ppt/comments/comment1.xml").async("text");
  assert.match(commentAuthorsXml, /<p:cmAuthor[^>]*name="Presentation Reviewer"/);
  assert.match(commentXml, /<p:cm[^>]*authorId="0"[^>]*idx="1"/);
  assert.match(commentXml, /Confirm the source before delivery\./);

  const modernCommentsDir = path.join(root, "modern-comments-workflow");
  const modernCommentsOutput = path.join(modernCommentsDir, "decision-review.pptx");
  const modernCommentsAudit = path.join(modernCommentsDir, "audit.json");
  const { createAndEditModernCommentThread } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-modern-comment-workflow.mjs"
  );
  const modernCommentsResult = await createAndEditModernCommentThread({
    outputPath: modernCommentsOutput,
    auditPath: modernCommentsAudit,
  });
  assert.equal(modernCommentsResult.audit.provider.actual, "office-kit");
  assert.equal(modernCommentsResult.audit.provider.silentFallback, false);
  assert.equal(modernCommentsResult.audit.operation.type, "fixed-topology-modern-comment-text-status-edit");
  assert.equal(modernCommentsResult.audit.operation.replyCount, 1);
  assert.equal(modernCommentsResult.audit.validation.fixedIdentityPreserved, true);
  assert.equal(modernCommentsResult.audit.validation.package.ok, true);
  assert.equal(modernCommentsResult.audit.validation.modelRender.ok, true);
  const modernCommentsRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(modernCommentsOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "decision-review.pptx",
  }));
  assert.equal(modernCommentsRoundTrip.commentFormat, "modern");
  const modernThread = modernCommentsRoundTrip.slides.getItem(0).comments.items[0];
  assert.equal(modernThread.comments[0].text, "Customer evidence confirmed for delivery.");
  assert.equal(modernThread.comments[1].text, "Recorded in the decision log.");
  assert.equal(modernThread.comments[0].author, "Review Owner");
  assert.equal(modernThread.comments[1].author, "Evidence Owner");
  assert.equal(modernThread.resolved, true);
  assert.equal(modernThread.nativeAnchor.type, "textRange");
  assert.match(modernCommentsRoundTrip.inspect({ kind: "comment" }).ndjson, /Customer evidence confirmed/);
  const modernCommentsZip = await JSZip.loadAsync(await fs.readFile(modernCommentsOutput));
  assert.ok(Object.keys(modernCommentsZip.files).some((name) => /^ppt\/comments\/modernComment\d*\.xml$/.test(name)));
  assert.ok(modernCommentsZip.file("ppt/authors.xml"));

  const modernCommentsCliOutput = path.join(modernCommentsDir, "decision-review-cli.pptx");
  const modernCommentsCliAudit = path.join(modernCommentsDir, "cli-audit.json");
  const modernCommentsCli = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/examples/officekit-modern-comment-workflow.mjs",
    modernCommentsCliOutput,
    modernCommentsCliAudit,
  ], { encoding: "utf8" });
  assert.equal(modernCommentsCli.status, 0, `modern-comment CLI failed\n${modernCommentsCli.stdout}\n${modernCommentsCli.stderr}`);
  assert.equal(JSON.parse(modernCommentsCli.stdout).threadId, "{11111111-1111-4111-8111-111111111111}");

  const coreEvidence = await runPresentationFixture(path.join(fixtureDir, "package-notes-comments.json"), {
    outputDir: path.join(root, "core-evidence"),
    nativeRender,
  });
  assert.equal(coreEvidence.qa.verify.ok, true);
  const evidenceSlide = coreEvidence.qa.presentation.slides.getItem(0);
  const evidenceCopy = itemByName(evidenceSlide.shapes.items, "evidence-copy");
  assert.equal(evidenceCopy.geometry, "textbox");
  assert.equal(evidenceCopy.text.paragraphs[1].bulletCharacter, "•");
  assert.deepEqual(evidenceCopy.text.paragraphs[3].runs[0].link, {
    uri: "https://www.ecma-international.org/publications-and-standards/standards/ecma-376/",
  });
  assert.equal(itemByName(evidenceSlide.charts.items, "evidence-pie").chartType, "pie");

  const titleNotesDir = path.join(root, "title-notes-workflow");
  const titleNotesInput = path.join(titleNotesDir, PPTX_TITLE_NOTES_FIXTURE.presentationName);
  const titleNotesOutput = path.join(titleNotesDir, "launch-review-updated.pptx");
  const titleNotesAudit = path.join(titleNotesDir, "audit.json");
  await generateOfficeInput("pptx-title-notes-review", titleNotesInput);
  const titleNotesSource = await fs.readFile(titleNotesInput);
  const { editPptxTitleAndNotes } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-title-notes-edit-workflow.mjs"
  );
  const titleNotesResult = await editPptxTitleAndNotes({
    inputPath: titleNotesInput,
    outputPath: titleNotesOutput,
    auditPath: titleNotesAudit,
    slideName: PPTX_TITLE_NOTES_FIXTURE.targetSlideName,
    titleShapeName: PPTX_TITLE_NOTES_FIXTURE.titleShapeName,
    expectedTitle: PPTX_TITLE_NOTES_FIXTURE.originalTitle,
    replacementTitle: PPTX_TITLE_NOTES_FIXTURE.replacementTitle,
    expectedNotes: PPTX_TITLE_NOTES_FIXTURE.originalNotes,
    replacementNotes: PPTX_TITLE_NOTES_FIXTURE.replacementNotes,
  });
  assert.equal(titleNotesResult.audit.provider.actual, "office-kit");
  assert.equal(titleNotesResult.audit.validation.reimport.ok, true);
  assert.equal(titleNotesResult.audit.validation.modelRender.renderer, "model-svg");
  assert.deepEqual(await fs.readFile(titleNotesInput), titleNotesSource);
  const titleNotesRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(titleNotesOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "launch-review-updated.pptx",
  }));
  assert.deepEqual(titleNotesRoundTrip.slides.items.map((slide) => slide.name), [
    PPTX_TITLE_NOTES_FIXTURE.targetSlideName,
    PPTX_TITLE_NOTES_FIXTURE.untouchedSlideName,
  ]);
  const titleNotesSlide = titleNotesRoundTrip.slides.getItem(0);
  assert.equal(itemByName(titleNotesSlide.shapes.items, PPTX_TITLE_NOTES_FIXTURE.titleShapeName).text.value, PPTX_TITLE_NOTES_FIXTURE.replacementTitle);
  assert.equal(titleNotesSlide.speakerNotes.text, PPTX_TITLE_NOTES_FIXTURE.replacementNotes);
  assert.deepEqual(titleNotesSlide.background, { fill: "#f1f5f9", mode: "solid" });

  const richNotesDir = path.join(root, "rich-notes-workflow");
  const richNotesInput = path.join(richNotesDir, PPTX_RICH_NOTES_FIXTURE.presentationName);
  const richNotesOutput = path.join(richNotesDir, "rich-notes-review-updated.pptx");
  const richNotesAudit = path.join(richNotesDir, "audit.json");
  await generateOfficeInput("pptx-rich-notes-review", richNotesInput);
  const richNotesSource = await fs.readFile(richNotesInput);
  const { editPptxRichSpeakerNotes } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-rich-speaker-notes-edit-workflow.mjs"
  );
  const richNotesResult = await editPptxRichSpeakerNotes({
    inputPath: richNotesInput,
    outputPath: richNotesOutput,
    auditPath: richNotesAudit,
  });
  assert.equal(richNotesResult.audit.provider.actual, "office-kit");
  assert.equal(richNotesResult.audit.operation.type, "title-and-rich-speaker-notes-run-edit");
  assert.equal(richNotesResult.audit.validation.reimport.richNotesFixedTopology, true);
  assert.deepEqual(await fs.readFile(richNotesInput), richNotesSource);
  const richNotesRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(richNotesOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "rich-notes-review-updated.pptx",
  }));
  const richNotesSlide = richNotesRoundTrip.slides.getItem(0);
  assert.equal(itemByName(richNotesSlide.shapes.items, PPTX_RICH_NOTES_FIXTURE.titleShapeName).text.value, PPTX_RICH_NOTES_FIXTURE.replacementTitle);
  assert.equal(richNotesSlide.speakerNotes.text, PPTX_RICH_NOTES_FIXTURE.replacementNotes);
  assert.deepEqual(richNotesSlide.speakerNotes.textFrame.paragraphs, [
    {
      runs: [
        { text: "Lead with ", style: { bold: true, fontSize: 18, fontFamily: "Aptos", color: "#0f172a" } },
        { text: PPTX_RICH_NOTES_FIXTURE.targetRun.replacementText, style: { bold: true, italic: false, fontSize: 18, color: "#0f766e" } },
      ],
      level: 0,
      bulletCharacter: "•",
      style: {},
    },
    {
      runs: [{ text: "Close with the accountable owner.", style: { fontSize: 16 } }],
      level: 0,
      autoNumber: { type: "arabicPeriod", startAt: 2 },
      style: {},
    },
  ]);

  const notesAddDir = path.join(root, "speaker-notes-add-workflow");
  const notesAddInput = path.join(notesAddDir, "speaker-notes-source.pptx");
  const notesAddOutput = path.join(notesAddDir, "speaker-notes-added.pptx");
  const notesAddAudit = path.join(notesAddDir, "audit.json");
  const notesText = "Lead with the verified evidence.\nClose with the requested decision.";
  const notesSourceDeck = Presentation.create();
  const notesTarget = notesSourceDeck.slides.add({ name: "Speaker notes target", background: { fill: "#F8FAFC" } });
  notesTarget.shapes.add({
    name: "notes-title",
    geometry: "textbox",
    text: "Visible slide content stays unchanged",
    position: { left: 96, top: 112, width: 1088, height: 96 },
  });
  const notesControl = notesSourceDeck.slides.add({ name: "Visual control", background: { fill: "#E0F2FE" } });
  notesControl.shapes.add({
    name: "control-title",
    geometry: "textbox",
    text: "Control slide",
    position: { left: 96, top: 112, width: 1088, height: 96 },
  });
  await fs.mkdir(notesAddDir, { recursive: true });
  await (await PresentationFile.exportPptx(notesSourceDeck)).save(notesAddInput);
  const notesSourceBytes = await fs.readFile(notesAddInput);
  const notesSourceImported = await PresentationFile.importPptx(new FileBlob(notesSourceBytes, {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "speaker-notes-source.pptx",
  }));
  assert.deepEqual(notesSourceImported.slides.getItem(0).speakerNotes.capability, {
    sourceBound: true,
    partPresent: false,
    editable: false,
    addable: true,
  });
  assert.match(notesSourceImported.inspect({ kind: "slide,notes" }).ndjson, /"notesCapability":\{"sourceBound":true,"partPresent":false,"editable":false,"addable":true\}/);
  const { addPptxSpeakerNotes } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-speaker-notes-add-workflow.mjs"
  );
  const notesAddResult = await addPptxSpeakerNotes({
    inputPath: notesAddInput,
    outputPath: notesAddOutput,
    auditPath: notesAddAudit,
    slideName: "Speaker notes target",
    notes: notesText,
  });
  assert.equal(notesAddResult.audit.provider.actual, "office-kit");
  assert.equal(notesAddResult.audit.operation.type, "source-bound-speaker-notes-add");
  assert.equal(notesAddResult.audit.precondition.capability.addable, true);
  assert.equal(notesAddResult.audit.validation.package.notesMasterPolicy, "created-canonical-shared-theme");
  assert.equal(notesAddResult.audit.validation.package.sourceHadNotesMaster, false);
  assert.equal(notesAddResult.audit.validation.package.notesSlideRelationshipCount, 2);
  assert.equal(notesAddResult.audit.validation.package.slideBackReferenceVerified, true);
  assert.equal(notesAddResult.audit.validation.modelRender.byteIdentical, true);
  assert.deepEqual(await fs.readFile(notesAddInput), notesSourceBytes);
  const notesRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(notesAddOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "speaker-notes-added.pptx",
  }));
  assert.equal(notesRoundTrip.slides.getItem(0).speakerNotes.text, notesText);
  assert.deepEqual(notesRoundTrip.slides.getItem(0).speakerNotes.capability, {
    sourceBound: true,
    partPresent: true,
    editable: true,
    addable: false,
  });
  assert.equal(notesRoundTrip.slides.getItem(1).speakerNotes.text, "");
  const notesZip = await JSZip.loadAsync(await fs.readFile(notesAddOutput));
  assert.equal(Object.keys(notesZip.files).filter((name) => /^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(name)).length, 1);
  assert.equal(Object.keys(notesZip.files).filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name)).length, 1);
  const notesBaselineDir = path.join(notesAddDir, "baselines");
  await verifyPresentationFile(notesAddInput, {
    outputDir: path.join(notesAddDir, "source-qa"),
    nativeRender,
    baselineDir: notesBaselineDir,
    writeBaseline: true,
  });
  const notesAddReview = await verifyPresentationFile(notesAddOutput, {
    outputDir: path.join(notesAddDir, "output-qa"),
    nativeRender,
    baselineDir: notesBaselineDir,
  });
  assert.ok(notesAddReview.modelRender.slides.every((slide) => slide.pixelDiff?.changed === false));
  if (nativeStatus.available) assert.ok(notesAddReview.nativeRender.pages.every((page) => page.pixelDiff?.changed === false));
  const notesRejectedOutput = path.join(notesAddDir, "speaker-notes-should-not-exist.pptx");
  const notesRejectedAudit = path.join(notesAddDir, "rejected-audit.json");
  await assert.rejects(
    () => addPptxSpeakerNotes({
      inputPath: notesAddOutput,
      outputPath: notesRejectedOutput,
      auditPath: notesRejectedAudit,
      slideName: "Speaker notes target",
      notes: "A second add must not masquerade as an edit.",
    }),
    /not an imported, notes-absent slide/i,
  );
  await assert.rejects(() => fs.access(notesRejectedOutput));
  await assert.rejects(() => fs.access(notesRejectedAudit));

  const legacyCommentAddDir = path.join(root, "legacy-comment-add-workflow");
  const legacyCommentAddInput = path.join(legacyCommentAddDir, "legacy-comment-source.pptx");
  const legacyCommentAddOutput = path.join(legacyCommentAddDir, "legacy-comment-added.pptx");
  const legacyCommentAddAudit = path.join(legacyCommentAddDir, "audit.json");
  const legacyCommentAddDeck = Presentation.create();
  const legacyCommentAddTarget = legacyCommentAddDeck.slides.add({ name: "Imported review target", background: { fill: "#F8FAFC" } });
  legacyCommentAddTarget.shapes.add({
    name: "review-title",
    geometry: "textbox",
    text: "Visible review content stays unchanged",
    position: { left: 96, top: 112, width: 1088, height: 96 },
  });
  const legacyCommentAddControl = legacyCommentAddDeck.slides.add({ name: "Imported review control", background: { fill: "#E0F2FE" } });
  legacyCommentAddControl.shapes.add({
    name: "control-title",
    geometry: "textbox",
    text: "Control slide",
    position: { left: 96, top: 112, width: 1088, height: 96 },
  });
  await fs.mkdir(legacyCommentAddDir, { recursive: true });
  await (await PresentationFile.exportPptx(legacyCommentAddDeck)).save(legacyCommentAddInput);
  const legacyCommentAddSourceBytes = await fs.readFile(legacyCommentAddInput);
  const legacyCommentAddImported = await PresentationFile.importPptx(new FileBlob(legacyCommentAddSourceBytes, {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "legacy-comment-source.pptx",
  }));
  assert.deepEqual(legacyCommentAddImported.slides.getItem(0).comments.capability, {
    sourceBound: true,
    format: "legacy",
    partPresent: false,
    editable: false,
    addable: true,
  });
  assert.match(legacyCommentAddImported.inspect({ kind: "slide" }).ndjson, /"commentsCapability":\{"sourceBound":true,"format":"legacy","partPresent":false,"editable":false,"addable":true\}/);
  const { addPptxLegacyReviewComment } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-legacy-comment-add-workflow.mjs"
  );
  const legacyCommentAddResult = await addPptxLegacyReviewComment({
    inputPath: legacyCommentAddInput,
    outputPath: legacyCommentAddOutput,
    auditPath: legacyCommentAddAudit,
    slideName: "Imported review target",
    text: "Confirm the imported evidence before delivery.",
    author: "Review Owner",
    created: "2026-07-20T03:04:05Z",
    position: { x: 360, y: 240, unit: "px" },
  });
  assert.equal(legacyCommentAddResult.audit.provider.actual, "office-kit");
  assert.equal(legacyCommentAddResult.audit.operation.type, "source-bound-legacy-comment-add");
  assert.equal(legacyCommentAddResult.audit.precondition.capability.addable, true);
  assert.deepEqual(legacyCommentAddResult.audit.validation.package.addedParts, [
    "ppt/commentAuthors.xml",
    "ppt/comments/comment1.xml",
  ]);
  assert.equal(legacyCommentAddResult.audit.validation.modelRender.byteIdentical, true);
  assert.deepEqual(await fs.readFile(legacyCommentAddInput), legacyCommentAddSourceBytes);
  const legacyCommentAddRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(legacyCommentAddOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "legacy-comment-added.pptx",
  }));
  assert.equal(legacyCommentAddRoundTrip.slides.getItem(0).comments.items[0].comments[0].text, "Confirm the imported evidence before delivery.");
  assert.deepEqual(legacyCommentAddRoundTrip.slides.getItem(0).comments.capability, {
    sourceBound: true,
    format: "legacy",
    partPresent: true,
    editable: true,
    addable: false,
  });
  assert.deepEqual(legacyCommentAddRoundTrip.slides.getItem(1).comments.capability, {
    sourceBound: true,
    format: "legacy",
    partPresent: false,
    editable: false,
    addable: false,
  });
  const legacyCommentAddZip = await JSZip.loadAsync(await fs.readFile(legacyCommentAddOutput));
  assert.equal(Object.keys(legacyCommentAddZip.files).filter((name) => /^ppt\/comments\/comment\d+\.xml$/.test(name)).length, 1);
  assert.ok(legacyCommentAddZip.file("ppt/commentAuthors.xml"));
  const legacyCommentAddBaselineDir = path.join(legacyCommentAddDir, "baselines");
  await verifyPresentationFile(legacyCommentAddInput, {
    outputDir: path.join(legacyCommentAddDir, "source-qa"),
    nativeRender,
    baselineDir: legacyCommentAddBaselineDir,
    writeBaseline: true,
  });
  const legacyCommentAddReview = await verifyPresentationFile(legacyCommentAddOutput, {
    outputDir: path.join(legacyCommentAddDir, "output-qa"),
    nativeRender,
    baselineDir: legacyCommentAddBaselineDir,
  });
  assert.ok(legacyCommentAddReview.modelRender.slides.every((slide) => slide.pixelDiff?.changed === false));
  if (nativeStatus.available) assert.ok(legacyCommentAddReview.nativeRender.pages.every((page) => page.pixelDiff?.changed === false));

  const legacyCommentEditOutput = path.join(legacyCommentAddDir, "legacy-comment-edited.pptx");
  const legacyCommentEditAudit = path.join(legacyCommentAddDir, "edit-audit.json");
  const legacyCommentEditSource = await fs.readFile(legacyCommentAddOutput);
  const legacyCommentEditSourceThread = legacyCommentAddRoundTrip.slides.getItem(0).comments.items[0];
  const { editPptxLegacyReviewComment } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-legacy-comment-edit-workflow.mjs"
  );
  const legacyCommentEditResult = await editPptxLegacyReviewComment({
    inputPath: legacyCommentAddOutput,
    outputPath: legacyCommentEditOutput,
    auditPath: legacyCommentEditAudit,
    slideName: "Imported review target",
    commentId: legacyCommentEditSourceThread.id,
    expectedText: "Confirm the imported evidence before delivery.",
    replacementText: "Confirm the imported evidence and record the delivery owner.",
  });
  assert.equal(legacyCommentEditResult.audit.operation.type, "source-bound-legacy-comment-text-edit");
  assert.equal(legacyCommentEditResult.audit.precondition.capability.editable, true);
  assert.deepEqual(legacyCommentEditResult.audit.validation.package.changedParts, ["ppt/comments/comment1.xml"]);
  assert.deepEqual(await fs.readFile(legacyCommentAddOutput), legacyCommentEditSource);
  const legacyCommentEditZip = await JSZip.loadAsync(await fs.readFile(legacyCommentEditOutput));
  for (const packagePath of ["ppt/commentAuthors.xml", "ppt/slides/slide1.xml", "ppt/slides/_rels/slide1.xml.rels", "[Content_Types].xml"]) {
    assert.deepEqual(
      await legacyCommentEditZip.file(packagePath).async("uint8array"),
      await legacyCommentAddZip.file(packagePath).async("uint8array"),
      `legacy comment edit must leave ${packagePath} byte-identical`,
    );
  }
  const legacyCommentEditRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(legacyCommentEditOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "legacy-comment-edited.pptx",
  }));
  const legacyCommentEditedThread = legacyCommentEditRoundTrip.slides.getItem(0).comments.items[0];
  assert.equal(legacyCommentEditedThread.comments[0].text, "Confirm the imported evidence and record the delivery owner.");
  assert.equal(legacyCommentEditedThread.comments[0].author, legacyCommentEditSourceThread.comments[0].author);
  assert.deepEqual(legacyCommentEditRoundTrip.slides.getItem(0).comments.capability, {
    sourceBound: true,
    format: "legacy",
    partPresent: true,
    editable: true,
    addable: false,
  });
  const legacyCommentEditReview = await verifyPresentationFile(legacyCommentEditOutput, {
    outputDir: path.join(legacyCommentAddDir, "edit-output-qa"),
    nativeRender,
    baselineDir: legacyCommentAddBaselineDir,
  });
  assert.ok(legacyCommentEditReview.modelRender.slides.every((slide) => slide.pixelDiff?.changed === false));
  if (nativeStatus.available) assert.ok(legacyCommentEditReview.nativeRender.pages.every((page) => page.pixelDiff?.changed === false));
  const legacyCommentEditRejectedOutput = path.join(legacyCommentAddDir, "legacy-comment-edit-should-not-exist.pptx");
  const legacyCommentEditRejectedAudit = path.join(legacyCommentAddDir, "edit-rejected-audit.json");
  await assert.rejects(
    () => editPptxLegacyReviewComment({
      inputPath: legacyCommentAddOutput,
      outputPath: legacyCommentEditRejectedOutput,
      auditPath: legacyCommentEditRejectedAudit,
      slideName: "Imported review target",
      commentId: legacyCommentEditSourceThread.id,
      expectedText: "A stale precondition must fail closed.",
      replacementText: "This must never publish.",
    }),
    /fixed root-text source-bound profile/i,
  );
  await assert.rejects(() => fs.access(legacyCommentEditRejectedOutput));
  await assert.rejects(() => fs.access(legacyCommentEditRejectedAudit));
  const legacyCommentRejectedOutput = path.join(legacyCommentAddDir, "legacy-comment-should-not-exist.pptx");
  const legacyCommentRejectedAudit = path.join(legacyCommentAddDir, "rejected-audit.json");
  await assert.rejects(
    () => addPptxLegacyReviewComment({
      inputPath: legacyCommentAddOutput,
      outputPath: legacyCommentRejectedOutput,
      auditPath: legacyCommentRejectedAudit,
      slideName: "Imported review target",
      text: "A second add must not masquerade as an edit.",
      author: "Review Owner",
      created: "2026-07-20T03:05:06Z",
      position: { x: 400, y: 260, unit: "px" },
    }),
    /not in an imported, comment-free presentation/i,
  );
  await assert.rejects(() => fs.access(legacyCommentRejectedOutput));
  await assert.rejects(() => fs.access(legacyCommentRejectedAudit));

  const slideNameDir = path.join(root, "slide-name-workflow");
  const slideNameInput = path.join(slideNameDir, PPTX_TITLE_NOTES_FIXTURE.presentationName);
  const slideNameOutput = path.join(slideNameDir, "launch-review-renamed.pptx");
  const slideNameAudit = path.join(slideNameDir, "audit.json");
  await generateOfficeInput("pptx-title-notes-review", slideNameInput);
  const slideNameSource = await fs.readFile(slideNameInput);
  const { editPptxSlideName } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-slide-name-edit-workflow.mjs"
  );
  const slideNameResult = await editPptxSlideName({
    inputPath: slideNameInput,
    outputPath: slideNameOutput,
    auditPath: slideNameAudit,
    expectedName: PPTX_TITLE_NOTES_FIXTURE.targetSlideName,
    replacementName: "Go decision: controlled rollout",
  });
  assert.equal(slideNameResult.audit.provider.actual, "office-kit");
  assert.equal(slideNameResult.audit.operation.nativeAttribute, "p:cSld/@name");
  assert.equal(slideNameResult.audit.validation.package.targetNameVerified, true);
  assert.equal(slideNameResult.audit.validation.package.targetPartMayBeCanonicalized, true);
  assert.equal(slideNameResult.audit.validation.package.nonTargetPartsByteIdentical, true);
  assert.equal(slideNameResult.audit.validation.modelRender.byteIdentical, true);
  assert.deepEqual(await fs.readFile(slideNameInput), slideNameSource);
  const slideNameRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(slideNameOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "launch-review-renamed.pptx",
  }));
  assert.deepEqual(slideNameRoundTrip.slides.items.map((slide) => slide.name), [
    "Go decision: controlled rollout",
    PPTX_TITLE_NOTES_FIXTURE.untouchedSlideName,
  ]);
  const slideNameTarget = slideNameRoundTrip.slides.getItem(0);
  assert.equal(itemByName(slideNameTarget.shapes.items, PPTX_TITLE_NOTES_FIXTURE.titleShapeName).text.value, PPTX_TITLE_NOTES_FIXTURE.originalTitle);
  assert.equal(slideNameTarget.speakerNotes.text, PPTX_TITLE_NOTES_FIXTURE.originalNotes);
  const missingOutput = path.join(slideNameDir, "should-not-exist.pptx");
  const missingAudit = path.join(slideNameDir, "should-not-exist.json");
  await assert.rejects(
    () => editPptxSlideName({
      inputPath: slideNameInput,
      outputPath: missingOutput,
      auditPath: missingAudit,
      expectedName: "Missing source slide",
      replacementName: "Never write this",
    }),
    /Expected exactly one imported slide named/,
  );
  assert.equal(await fs.access(missingOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(missingAudit).then(() => true, () => false), false);
  const slideNameCliOutput = path.join(slideNameDir, "launch-review-cli-renamed.pptx");
  const slideNameCliAudit = path.join(slideNameDir, "cli-audit.json");
  const slideNameCli = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/examples/officekit-slide-name-edit-workflow.mjs",
    slideNameInput,
    slideNameCliOutput,
    slideNameCliAudit,
    PPTX_TITLE_NOTES_FIXTURE.targetSlideName,
    "Go decision: CLI rollout",
  ], { encoding: "utf8" });
  assert.equal(slideNameCli.status, 0, `slide-name CLI failed\n${slideNameCli.stdout}\n${slideNameCli.stderr}`);
  assert.equal(JSON.parse(slideNameCli.stdout).sourcePart, "ppt/slides/slide1.xml");

  const viewPropertiesDir = path.join(root, "view-properties-workflow");
  const viewPropertiesInput = path.join(viewPropertiesDir, "view-properties-source.pptx");
  const viewPropertiesOutput = path.join(viewPropertiesDir, "view-properties-updated.pptx");
  const viewPropertiesAudit = path.join(viewPropertiesDir, "audit.json");
  const viewPropertiesDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
  viewPropertiesDeck.slides.add({ name: "View properties target" }).shapes.add({
    name: "view-properties-title",
    text: "This slide must remain visually unchanged",
    // Leave enough height for the intentional two-line fixture title so the
    // workflow's visual verification exercises the view-properties change,
    // rather than accepting a pre-existing text-overflow warning.
    position: { left: 48, top: 72, width: 520, height: 80 },
  });
  const viewPropertiesBase = await PresentationFile.exportPptx(viewPropertiesDeck);
  const viewPropertiesBaseZip = await JSZip.loadAsync(viewPropertiesBase.bytes);
  const viewPropertiesRelationships = await viewPropertiesBaseZip.file("ppt/_rels/presentation.xml.rels").async("text");
  const viewPropertiesXml = '<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" lastView="sldView"><p:slideViewPr><p:cSldViewPr snapToGrid="1" snapToObjects="0" showGuides="1"><p:cViewPr varScale="1"><p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale><p:origin x="0" y="0"/></p:cViewPr><p:guideLst><p:guide orient="horz" pos="2160"/><p:guide orient="vert" pos="2880"/></p:guideLst></p:cSldViewPr></p:slideViewPr><p:gridSpacing cx="72008" cy="91440"/></p:viewPr>';
  const viewPropertiesSource = await PresentationFile.patchPptx(viewPropertiesBase, [
    {
      path: "ppt/_rels/presentation.xml.rels",
      xml: viewPropertiesRelationships.replace("</Relationships>", '<Relationship Id="rIdViewProperties" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/></Relationships>'),
    },
    {
      path: "ppt/viewProps.xml",
      xml: viewPropertiesXml,
      contentType: "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml",
    },
  ]);
  await fs.mkdir(viewPropertiesDir, { recursive: true });
  await viewPropertiesSource.save(viewPropertiesInput);
  const viewPropertiesSourceBytes = await fs.readFile(viewPropertiesInput);
  const { editPptxViewProperties } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-view-properties-edit-workflow.mjs"
  );
  const viewPropertiesResult = await editPptxViewProperties({
    inputPath: viewPropertiesInput,
    outputPath: viewPropertiesOutput,
    auditPath: viewPropertiesAudit,
    patch: {
      gridSpacingCxEmu: 72_009,
      gridSpacingCyEmu: 91_441,
      slideViewSnapToGrid: false,
      slideViewSnapToObjects: true,
      slideGuides: [
        { orientation: "horizontal", position: 2_161 },
        { orientation: "vertical", position: 2_881 },
      ],
    },
  });
  assert.equal(viewPropertiesResult.audit.operation.type, "source-bound-view-properties-edit");
  assert.equal(viewPropertiesResult.audit.validation.package.onlyViewPropertiesPartChanged, true);
  assert.equal(viewPropertiesResult.audit.validation.package.guideVisibilityPreserved, true);
  assert.equal(viewPropertiesResult.audit.validation.reimport.fixedTopologyEditable, true);
  assert.equal(viewPropertiesResult.audit.validation.modelRender.byteIdentical, true);
  assert.deepEqual(await fs.readFile(viewPropertiesInput), viewPropertiesSourceBytes);
  const viewPropertiesRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(viewPropertiesOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "view-properties-updated.pptx",
  }));
  assert.equal(viewPropertiesRoundTrip.view.gridSpacingCxEmu, 72_009);
  assert.equal(viewPropertiesRoundTrip.view.gridSpacingCyEmu, 91_441);
  assert.equal(viewPropertiesRoundTrip.view.slideViewSnapToGrid, false);
  assert.equal(viewPropertiesRoundTrip.view.slideViewSnapToObjects, true);
  assert.deepEqual(viewPropertiesRoundTrip.view.slideGuides, [
    { orientation: "horizontal", position: 2_161 },
    { orientation: "vertical", position: 2_881 },
  ]);
  const viewPropertiesRejectedOutput = path.join(viewPropertiesDir, "view-properties-should-not-exist.pptx");
  const viewPropertiesRejectedAudit = path.join(viewPropertiesDir, "rejected-audit.json");
  await assert.rejects(
    () => editPptxViewProperties({
      inputPath: viewPropertiesInput,
      outputPath: viewPropertiesRejectedOutput,
      auditPath: viewPropertiesRejectedAudit,
      patch: { slideGuides: [{ orientation: "horizontal", position: 2161 }] },
    }),
    /guide count, order, and orientation are source-bound/,
  );
  assert.equal(await fs.access(viewPropertiesRejectedOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(viewPropertiesRejectedAudit).then(() => true, () => false), false);
  const viewPropertiesCliOutput = path.join(viewPropertiesDir, "view-properties-cli.pptx");
  const viewPropertiesCliAudit = path.join(viewPropertiesDir, "cli-audit.json");
  const viewPropertiesCli = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/examples/officekit-view-properties-edit-workflow.mjs",
    viewPropertiesInput,
    viewPropertiesCliOutput,
    viewPropertiesCliAudit,
    JSON.stringify({ gridSpacingCxEmu: 72_010 }),
  ], { encoding: "utf8" });
  assert.equal(viewPropertiesCli.status, 0, `view-properties CLI failed\n${viewPropertiesCli.stdout}\n${viewPropertiesCli.stderr}`);
  assert.equal(JSON.parse(viewPropertiesCli.stdout).targetPart, "ppt/viewProps.xml");

  const transitionEditDir = path.join(root, "transition-edit-workflow");
  const transitionEditInput = path.join(transitionEditDir, "transition-source.pptx");
  const transitionEditOutput = path.join(transitionEditDir, "transition-updated.pptx");
  const transitionEditAudit = path.join(transitionEditDir, "audit.json");
  const transitionEditDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
  const transitionEditTarget = transitionEditDeck.slides.add({ name: "Transition review target" });
  transitionEditTarget.shapes.add({
    name: "transition-review-title",
    position: { left: 48, top: 72, width: 520, height: 72 },
    text: "Visible slide content must remain stable",
  });
  transitionEditTarget.setTransition({ effect: "fade", throughBlack: true, speed: "medium", durationMs: 700, advanceOnClick: true, advanceAfterMs: 1_200 });
  const transitionEditUntouched = transitionEditDeck.slides.add({ name: "Transition appendix" });
  transitionEditUntouched.shapes.add({
    name: "transition-appendix-title",
    position: { left: 48, top: 72, width: 520, height: 72 },
    text: "Appendix slide remains byte-stable",
  });
  transitionEditUntouched.setTransition({ effect: "wheel", spokes: 8, speed: "fast", advanceOnClick: false });
  await fs.mkdir(transitionEditDir, { recursive: true });
  await (await PresentationFile.exportPptx(transitionEditDeck)).save(transitionEditInput);
  const transitionEditSource = await fs.readFile(transitionEditInput);
  const { editPptxTransition } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-transition-edit-workflow.mjs"
  );
  const transitionEditResult = await editPptxTransition({
    inputPath: transitionEditInput,
    outputPath: transitionEditOutput,
    auditPath: transitionEditAudit,
    slideName: "Transition review target",
    expectedTransition: { effect: "fade", throughBlack: true, speed: "medium", durationMs: 700, advanceOnClick: true, advanceAfterMs: 1_200 },
    replacementTransition: { effect: "split", orientation: "horizontal", direction: "in", speed: "slow", durationMs: 1_100, advanceOnClick: false },
  });
  assert.equal(transitionEditResult.audit.operation.type, "source-bound-transition-edit");
  assert.equal(transitionEditResult.audit.operation.partPath, "ppt/slides/slide1.xml");
  assert.equal(transitionEditResult.audit.validation.package.onlyTargetSlidePartChanged, true);
  assert.equal(transitionEditResult.audit.validation.package.nonTargetPartsByteIdentical, true);
  assert.equal(transitionEditResult.audit.validation.reimport.replacementSemanticsRetained, true);
  assert.equal(transitionEditResult.audit.validation.nonTransitionSemantics.stable, true);
  assert.equal(transitionEditResult.audit.validation.modelRender.byteIdentical, true);
  assert.deepEqual(await fs.readFile(transitionEditInput), transitionEditSource);
  const transitionEditOutputZip = await JSZip.loadAsync(await fs.readFile(transitionEditOutput));
  const transitionEditSourceZip = await JSZip.loadAsync(transitionEditSource);
  assert.notDeepEqual(
    await transitionEditOutputZip.file("ppt/slides/slide1.xml").async("uint8array"),
    await transitionEditSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  );
  assert.deepEqual(
    await transitionEditOutputZip.file("ppt/slides/slide2.xml").async("uint8array"),
    await transitionEditSourceZip.file("ppt/slides/slide2.xml").async("uint8array"),
  );
  const transitionEditRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(transitionEditOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "transition-updated.pptx",
  }));
  assert.deepEqual(transitionEditRoundTrip.slides.getItem(0).transition.toJSON(), {
    effect: "split",
    direction: "in",
    orientation: "horizontal",
    speed: "slow",
    durationMs: 1_100,
    advanceOnClick: false,
  });
  assert.deepEqual(transitionEditRoundTrip.slides.getItem(1).transition.toJSON(), {
    effect: "wheel",
    spokes: 8,
    speed: "fast",
    advanceOnClick: false,
  });
  assert.deepEqual(transitionEditRoundTrip.slides.getItem(0).transition.capability, {
    sourceBound: true,
    partPresent: true,
    editable: true,
    addable: false,
  });
  const transitionEditBaselineDir = path.join(transitionEditDir, "baselines");
  await verifyPresentationFile(transitionEditInput, {
    outputDir: path.join(transitionEditDir, "source-qa"),
    nativeRender,
    baselineDir: transitionEditBaselineDir,
    writeBaseline: true,
  });
  const transitionEditReview = await verifyPresentationFile(transitionEditOutput, {
    outputDir: path.join(transitionEditDir, "output-qa"),
    nativeRender,
    baselineDir: transitionEditBaselineDir,
  });
  assert.ok(transitionEditReview.modelRender.slides.every((slide) => slide.pixelDiff?.changed === false));
  if (nativeStatus.available) assert.ok(transitionEditReview.nativeRender.pages.every((page) => page.pixelDiff?.changed === false));
  const transitionEditRejectedOutput = path.join(transitionEditDir, "transition-should-not-exist.pptx");
  const transitionEditRejectedAudit = path.join(transitionEditDir, "rejected-audit.json");
  await assert.rejects(
    () => editPptxTransition({
      inputPath: transitionEditInput,
      outputPath: transitionEditRejectedOutput,
      auditPath: transitionEditRejectedAudit,
      slideName: "Transition review target",
      expectedTransition: { effect: "fade", throughBlack: true, speed: "medium", durationMs: 700, advanceOnClick: true, advanceAfterMs: 1_201 },
      replacementTransition: { effect: "split", orientation: "horizontal", direction: "in", speed: "slow", durationMs: 1_100, advanceOnClick: false },
    }),
    /does not match expectedTransition/i,
  );
  assert.equal(await fs.access(transitionEditRejectedOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(transitionEditRejectedAudit).then(() => true, () => false), false);
  const transitionEditCollisionOutput = path.join(transitionEditDir, "transition-existing-output.pptx");
  const transitionEditCollisionAudit = path.join(transitionEditDir, "collision-audit.json");
  const collisionSentinel = Buffer.from("preserve-existing-output");
  await fs.writeFile(transitionEditCollisionOutput, collisionSentinel);
  await assert.rejects(
    () => editPptxTransition({
      inputPath: transitionEditInput,
      outputPath: transitionEditCollisionOutput,
      auditPath: transitionEditCollisionAudit,
      slideName: "Transition review target",
      expectedTransition: { effect: "fade", throughBlack: true, speed: "medium", durationMs: 700, advanceOnClick: true, advanceAfterMs: 1_200 },
      replacementTransition: { effect: "split", orientation: "horizontal", direction: "in", speed: "slow", durationMs: 1_100, advanceOnClick: false },
    }),
    /outputPath already exists; refusing to overwrite/i,
  );
  assert.deepEqual(await fs.readFile(transitionEditCollisionOutput), collisionSentinel);
  assert.equal(await fs.access(transitionEditCollisionAudit).then(() => true, () => false), false);
  const transitionEditCliOutput = path.join(transitionEditDir, "transition-cli.pptx");
  const transitionEditCliAudit = path.join(transitionEditDir, "cli-audit.json");
  const transitionEditCli = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/examples/officekit-transition-edit-workflow.mjs",
    transitionEditInput,
    transitionEditCliOutput,
    transitionEditCliAudit,
    "Transition review target",
    JSON.stringify({ effect: "fade", throughBlack: true, speed: "medium", durationMs: 700, advanceOnClick: true, advanceAfterMs: 1_200 }),
    JSON.stringify({ effect: "cover", direction: "rightUp", speed: "fast", durationMs: 1_500, advanceOnClick: true, advanceAfterMs: 2_000 }),
  ], { encoding: "utf8" });
  assert.equal(transitionEditCli.status, 0, `transition-edit CLI failed\n${transitionEditCli.stdout}\n${transitionEditCli.stderr}`);
  assert.equal(JSON.parse(transitionEditCli.stdout).targetPart, "ppt/slides/slide1.xml");

  const sectionRenameDir = path.join(root, "section-rename-workflow");
  const sectionRenameInput = path.join(sectionRenameDir, "section-source.pptx");
  const sectionRenameOutput = path.join(sectionRenameDir, "section-updated.pptx");
  const sectionRenameAudit = path.join(sectionRenameDir, "audit.json");
  const sectionRenameDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
  const sectionRenameSlides = [
    ["Section opening", "#DBEAFE"],
    ["Section evidence", "#DCFCE7"],
    ["Section decision", "#FEF3C7"],
  ].map(([name, fill], index) => {
    const slide = sectionRenameDeck.slides.add({ name, background: { fill } });
    slide.shapes.add({
      name: `section-title-${index + 1}`,
      position: { left: 48, top: 72, width: 520, height: 72 },
      text: `${index + 1}. ${name}`,
    });
    return slide;
  });
  sectionRenameDeck.sections.add({
    name: "Context",
    nativeId: "{01F07B81-39E6-4BBB-9B89-66EA253FBD29}",
    slides: sectionRenameSlides.slice(0, 2),
  });
  sectionRenameDeck.sections.add({
    name: "Decision",
    nativeId: "{1FEF2C88-0CF2-4176-BA81-0DE6FD9D1274}",
    slides: sectionRenameSlides.slice(2),
  });
  await fs.mkdir(sectionRenameDir, { recursive: true });
  await (await PresentationFile.exportPptx(sectionRenameDeck)).save(sectionRenameInput);
  const sectionRenameSource = await fs.readFile(sectionRenameInput);
  const { renamePptxSection } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-section-rename-workflow.mjs"
  );
  const sectionRenameResult = await renamePptxSection({
    inputPath: sectionRenameInput,
    outputPath: sectionRenameOutput,
    auditPath: sectionRenameAudit,
    expectedName: "Context",
    replacementName: "Background",
  });
  assert.equal(sectionRenameResult.audit.operation.type, "source-bound-section-name-edit");
  assert.equal(sectionRenameResult.audit.operation.sectionId, "section/1");
  assert.equal(sectionRenameResult.audit.operation.nativeId, "{01F07B81-39E6-4BBB-9B89-66EA253FBD29}");
  assert.deepEqual(sectionRenameResult.audit.operation.orderedSlideIds, ["presentation/slide/1", "presentation/slide/2"]);
  assert.equal(sectionRenameResult.audit.validation.package.targetPart, "ppt/presentation.xml");
  assert.equal(sectionRenameResult.audit.validation.package.onlyPresentationPartChanged, true);
  assert.equal(sectionRenameResult.audit.validation.package.nonTargetPartsByteIdentical, true);
  assert.equal(sectionRenameResult.audit.validation.reimport.exactFixedTopologyRetained, true);
  assert.equal(sectionRenameResult.audit.validation.nonSectionSemantics.stable, true);
  assert.equal(sectionRenameResult.audit.validation.modelRender.byteIdentical, true);
  assert.deepEqual(await fs.readFile(sectionRenameInput), sectionRenameSource);
  const sectionRenameSourceZip = await JSZip.loadAsync(sectionRenameSource);
  const sectionRenameOutputZip = await JSZip.loadAsync(await fs.readFile(sectionRenameOutput));
  assert.notDeepEqual(
    await sectionRenameSourceZip.file("ppt/presentation.xml").async("uint8array"),
    await sectionRenameOutputZip.file("ppt/presentation.xml").async("uint8array"),
  );
  for (const partPath of Object.keys(sectionRenameSourceZip.files)) {
    if (sectionRenameSourceZip.files[partPath].dir || partPath === "ppt/presentation.xml") continue;
    assert.deepEqual(
      await sectionRenameOutputZip.file(partPath).async("uint8array"),
      await sectionRenameSourceZip.file(partPath).async("uint8array"),
      `section rename changed non-target part ${partPath}`,
    );
  }
  const sectionRenameRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(sectionRenameOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "section-updated.pptx",
  }));
  assert.deepEqual(sectionRenameRoundTrip.sections.items.map((section) => section.toJSON()), [
    {
      id: "section/1",
      name: "Background",
      nativeId: "{01F07B81-39E6-4BBB-9B89-66EA253FBD29}",
      slideIds: ["presentation/slide/1", "presentation/slide/2"],
    },
    {
      id: "section/2",
      name: "Decision",
      nativeId: "{1FEF2C88-0CF2-4176-BA81-0DE6FD9D1274}",
      slideIds: ["presentation/slide/3"],
    },
  ]);
  const sectionRenameBaselineDir = path.join(sectionRenameDir, "baselines");
  await verifyPresentationFile(sectionRenameInput, {
    outputDir: path.join(sectionRenameDir, "source-qa"),
    nativeRender,
    baselineDir: sectionRenameBaselineDir,
    writeBaseline: true,
  });
  const sectionRenameReview = await verifyPresentationFile(sectionRenameOutput, {
    outputDir: path.join(sectionRenameDir, "output-qa"),
    nativeRender,
    baselineDir: sectionRenameBaselineDir,
  });
  assert.ok(sectionRenameReview.modelRender.slides.every((slide) => slide.pixelDiff?.changed === false));
  if (nativeStatus.available) assert.ok(sectionRenameReview.nativeRender.pages.every((page) => page.pixelDiff?.changed === false));
  const sectionRenameRejectedOutput = path.join(sectionRenameDir, "section-should-not-exist.pptx");
  const sectionRenameRejectedAudit = path.join(sectionRenameDir, "rejected-audit.json");
  await assert.rejects(
    () => renamePptxSection({
      inputPath: sectionRenameInput,
      outputPath: sectionRenameRejectedOutput,
      auditPath: sectionRenameRejectedAudit,
      expectedName: "Stale context",
      replacementName: "Background",
    }),
    /Expected exactly one imported PowerPoint section named/i,
  );
  assert.equal(await fs.access(sectionRenameRejectedOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(sectionRenameRejectedAudit).then(() => true, () => false), false);
  const sectionRenameConflictOutput = path.join(sectionRenameDir, "section-conflict-output.pptx");
  const sectionRenameConflictAudit = path.join(sectionRenameDir, "section-conflict-audit.json");
  await assert.rejects(
    () => renamePptxSection({
      inputPath: sectionRenameInput,
      outputPath: sectionRenameConflictOutput,
      auditPath: sectionRenameConflictAudit,
      expectedName: "Context",
      replacementName: "Decision",
    }),
    /conflicts case-insensitively/i,
  );
  assert.equal(await fs.access(sectionRenameConflictOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(sectionRenameConflictAudit).then(() => true, () => false), false);
  const sectionRenameOpaqueInput = path.join(sectionRenameDir, "section-opaque-source.pptx");
  const sectionRenameOpaqueOutput = path.join(sectionRenameDir, "section-opaque-output.pptx");
  const sectionRenameOpaqueAudit = path.join(sectionRenameDir, "section-opaque-audit.json");
  const sectionRenameOpaqueZip = await JSZip.loadAsync(sectionRenameSource);
  const sectionRenamePresentationXml = await sectionRenameOpaqueZip.file("ppt/presentation.xml").async("text");
  sectionRenameOpaqueZip.file("ppt/presentation.xml", sectionRenamePresentationXml.replace('name="Context"', 'name="Context" unmodeled="true"'));
  await fs.writeFile(sectionRenameOpaqueInput, await sectionRenameOpaqueZip.generateAsync({ type: "nodebuffer" }));
  await assert.rejects(
    () => renamePptxSection({
      inputPath: sectionRenameOpaqueInput,
      outputPath: sectionRenameOpaqueOutput,
      auditPath: sectionRenameOpaqueAudit,
      expectedName: "Context",
      replacementName: "Background",
    }),
    /no semantic canonical PowerPoint sections/i,
  );
  assert.equal(await fs.access(sectionRenameOpaqueOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(sectionRenameOpaqueAudit).then(() => true, () => false), false);
  const sectionRenameCollisionOutput = path.join(sectionRenameDir, "section-existing-output.pptx");
  const sectionRenameCollisionAudit = path.join(sectionRenameDir, "collision-audit.json");
  const sectionRenameCollisionSentinel = Buffer.from("preserve-existing-section-output");
  await fs.writeFile(sectionRenameCollisionOutput, sectionRenameCollisionSentinel);
  await assert.rejects(
    () => renamePptxSection({
      inputPath: sectionRenameInput,
      outputPath: sectionRenameCollisionOutput,
      auditPath: sectionRenameCollisionAudit,
      expectedName: "Context",
      replacementName: "Background",
    }),
    /outputPath already exists; refusing to overwrite/i,
  );
  assert.deepEqual(await fs.readFile(sectionRenameCollisionOutput), sectionRenameCollisionSentinel);
  assert.equal(await fs.access(sectionRenameCollisionAudit).then(() => true, () => false), false);
  const sectionRenameCliOutput = path.join(sectionRenameDir, "section-cli.pptx");
  const sectionRenameCliAudit = path.join(sectionRenameDir, "cli-audit.json");
  const sectionRenameCli = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/examples/officekit-section-rename-workflow.mjs",
    sectionRenameInput,
    sectionRenameCliOutput,
    sectionRenameCliAudit,
    "Context",
    "Narrative context",
  ], { encoding: "utf8" });
  assert.equal(sectionRenameCli.status, 0, `section-rename CLI failed\n${sectionRenameCli.stdout}\n${sectionRenameCli.stderr}`);
  assert.equal(JSON.parse(sectionRenameCli.stdout).targetPart, "ppt/presentation.xml");

  const sectionBoundaryDir = path.join(root, "section-boundary-workflow");
  const sectionBoundaryInput = path.join(sectionBoundaryDir, "section-boundary-source.pptx");
  const sectionBoundaryOutput = path.join(sectionBoundaryDir, "section-boundary-updated.pptx");
  const sectionBoundaryAudit = path.join(sectionBoundaryDir, "audit.json");
  const sectionBoundaryDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
  const sectionBoundarySlides = [
    ["Boundary opening", "#DBEAFE"],
    ["Boundary evidence", "#DCFCE7"],
    ["Boundary decision", "#FEF3C7"],
    ["Boundary appendix", "#FCE7F3"],
  ].map(([name, fill], index) => {
    const slide = sectionBoundaryDeck.slides.add({ name, background: { fill } });
    slide.shapes.add({
      name: `section-boundary-title-${index + 1}`,
      position: { left: 48, top: 72, width: 520, height: 72 },
      text: `${index + 1}. ${name}`,
    });
    return slide;
  });
  sectionBoundaryDeck.sections.add({
    name: "Context",
    nativeId: "{01F07B81-39E6-4BBB-9B89-66EA253FBD29}",
    slides: sectionBoundarySlides.slice(0, 2),
  });
  sectionBoundaryDeck.sections.add({
    name: "Decision",
    nativeId: "{1FEF2C88-0CF2-4176-BA81-0DE6FD9D1274}",
    slides: sectionBoundarySlides.slice(2, 3),
  });
  sectionBoundaryDeck.sections.add({
    name: "Appendix",
    nativeId: "{2E92C0F3-07D0-4D22-8AC3-55C9651C42B1}",
    slides: sectionBoundarySlides.slice(3),
  });
  const expectedSectionPartition = [
    {
      id: "section/1",
      name: "Context",
      nativeId: "{01F07B81-39E6-4BBB-9B89-66EA253FBD29}",
      slideIds: ["presentation/slide/1", "presentation/slide/2"],
    },
    {
      id: "section/2",
      name: "Decision",
      nativeId: "{1FEF2C88-0CF2-4176-BA81-0DE6FD9D1274}",
      slideIds: ["presentation/slide/3"],
    },
    {
      id: "section/3",
      name: "Appendix",
      nativeId: "{2E92C0F3-07D0-4D22-8AC3-55C9651C42B1}",
      slideIds: ["presentation/slide/4"],
    },
  ];
  const replacementSectionPartition = [
    {
      ...expectedSectionPartition[0],
      slideIds: ["presentation/slide/1"],
    },
    {
      ...expectedSectionPartition[1],
      slideIds: ["presentation/slide/2", "presentation/slide/3"],
    },
    expectedSectionPartition[2],
  ];
  await fs.mkdir(sectionBoundaryDir, { recursive: true });
  await (await PresentationFile.exportPptx(sectionBoundaryDeck)).save(sectionBoundaryInput);
  const sectionBoundarySource = await fs.readFile(sectionBoundaryInput);
  const { replacePptxSectionPartition } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-section-boundary-edit-workflow.mjs"
  );
  const sectionBoundaryResult = await replacePptxSectionPartition({
    inputPath: sectionBoundaryInput,
    outputPath: sectionBoundaryOutput,
    auditPath: sectionBoundaryAudit,
    expectedSections: expectedSectionPartition,
    replacementSections: replacementSectionPartition,
  });
  assert.equal(sectionBoundaryResult.audit.operation.type, "source-bound-section-boundary-edit");
  assert.equal(sectionBoundaryResult.audit.operation.fixedSectionCount, 3);
  assert.equal(sectionBoundaryResult.audit.operation.changedSections.length, 2);
  assert.deepEqual(sectionBoundaryResult.audit.operation.replacementSections, replacementSectionPartition);
  assert.equal(sectionBoundaryResult.audit.validation.package.targetPart, "ppt/presentation.xml");
  assert.equal(sectionBoundaryResult.audit.validation.package.onlyPresentationPartChanged, true);
  assert.equal(sectionBoundaryResult.audit.validation.package.nonTargetPartsByteIdentical, true);
  assert.equal(sectionBoundaryResult.audit.validation.reimport.exactFixedIdentityPartitionRetained, true);
  assert.equal(sectionBoundaryResult.audit.validation.nonSectionSemantics.stable, true);
  assert.equal(sectionBoundaryResult.audit.validation.modelRender.byteIdentical, true);
  assert.deepEqual(await fs.readFile(sectionBoundaryInput), sectionBoundarySource);
  const sectionBoundarySourceZip = await JSZip.loadAsync(sectionBoundarySource);
  const sectionBoundaryOutputZip = await JSZip.loadAsync(await fs.readFile(sectionBoundaryOutput));
  assert.notDeepEqual(
    await sectionBoundarySourceZip.file("ppt/presentation.xml").async("uint8array"),
    await sectionBoundaryOutputZip.file("ppt/presentation.xml").async("uint8array"),
  );
  for (const partPath of Object.keys(sectionBoundarySourceZip.files)) {
    if (sectionBoundarySourceZip.files[partPath].dir || partPath === "ppt/presentation.xml") continue;
    assert.deepEqual(
      await sectionBoundaryOutputZip.file(partPath).async("uint8array"),
      await sectionBoundarySourceZip.file(partPath).async("uint8array"),
      `section boundary edit changed non-target part ${partPath}`,
    );
  }
  const sectionBoundaryRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(sectionBoundaryOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "section-boundary-updated.pptx",
  }));
  assert.deepEqual(sectionBoundaryRoundTrip.sections.items.map((section) => section.toJSON()), replacementSectionPartition);
  const sectionBoundaryBaselineDir = path.join(sectionBoundaryDir, "baselines");
  await verifyPresentationFile(sectionBoundaryInput, {
    outputDir: path.join(sectionBoundaryDir, "source-qa"),
    nativeRender,
    baselineDir: sectionBoundaryBaselineDir,
    writeBaseline: true,
  });
  const sectionBoundaryReview = await verifyPresentationFile(sectionBoundaryOutput, {
    outputDir: path.join(sectionBoundaryDir, "output-qa"),
    nativeRender,
    baselineDir: sectionBoundaryBaselineDir,
  });
  assert.ok(sectionBoundaryReview.modelRender.slides.every((slide) => slide.pixelDiff?.changed === false));
  if (nativeStatus.available) assert.ok(sectionBoundaryReview.nativeRender.pages.every((page) => page.pixelDiff?.changed === false));
  const sectionBoundaryStaleOutput = path.join(sectionBoundaryDir, "section-boundary-stale-output.pptx");
  const sectionBoundaryStaleAudit = path.join(sectionBoundaryDir, "stale-audit.json");
  await assert.rejects(
    () => replacePptxSectionPartition({
      inputPath: sectionBoundaryInput,
      outputPath: sectionBoundaryStaleOutput,
      auditPath: sectionBoundaryStaleAudit,
      expectedSections: [{ ...expectedSectionPartition[0], name: "Stale context" }, ...expectedSectionPartition.slice(1)],
      replacementSections: replacementSectionPartition,
    }),
    /expectedSections does not exactly match/i,
  );
  assert.equal(await fs.access(sectionBoundaryStaleOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(sectionBoundaryStaleAudit).then(() => true, () => false), false);
  const sectionBoundaryNoncanonicalOutput = path.join(sectionBoundaryDir, "section-boundary-noncanonical-output.pptx");
  const sectionBoundaryNoncanonicalAudit = path.join(sectionBoundaryDir, "noncanonical-audit.json");
  await assert.rejects(
    () => replacePptxSectionPartition({
      inputPath: sectionBoundaryInput,
      outputPath: sectionBoundaryNoncanonicalOutput,
      auditPath: sectionBoundaryNoncanonicalAudit,
      expectedSections: [{ ...expectedSectionPartition[0], name: "Context " }, ...expectedSectionPartition.slice(1)],
      replacementSections: replacementSectionPartition,
    }),
    /without leading or trailing whitespace/i,
  );
  assert.equal(await fs.access(sectionBoundaryNoncanonicalOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(sectionBoundaryNoncanonicalAudit).then(() => true, () => false), false);
  const sectionBoundaryPartialOutput = path.join(sectionBoundaryDir, "section-boundary-partial-output.pptx");
  const sectionBoundaryPartialAudit = path.join(sectionBoundaryDir, "partial-audit.json");
  await assert.rejects(
    () => replacePptxSectionPartition({
      inputPath: sectionBoundaryInput,
      outputPath: sectionBoundaryPartialOutput,
      auditPath: sectionBoundaryPartialAudit,
      expectedSections: expectedSectionPartition,
      replacementSections: replacementSectionPartition.slice(0, 2),
    }),
    /must each list every fixed source section/i,
  );
  assert.equal(await fs.access(sectionBoundaryPartialOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(sectionBoundaryPartialAudit).then(() => true, () => false), false);
  const sectionBoundaryRenameOutput = path.join(sectionBoundaryDir, "section-boundary-rename-output.pptx");
  const sectionBoundaryRenameAudit = path.join(sectionBoundaryDir, "rename-audit.json");
  await assert.rejects(
    () => replacePptxSectionPartition({
      inputPath: sectionBoundaryInput,
      outputPath: sectionBoundaryRenameOutput,
      auditPath: sectionBoundaryRenameAudit,
      expectedSections: expectedSectionPartition,
      replacementSections: [{ ...replacementSectionPartition[0], name: "Background" }, ...replacementSectionPartition.slice(1)],
    }),
    /must preserve the source section ID, name, and native GUID/i,
  );
  assert.equal(await fs.access(sectionBoundaryRenameOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(sectionBoundaryRenameAudit).then(() => true, () => false), false);
  const sectionBoundaryOrderOutput = path.join(sectionBoundaryDir, "section-boundary-order-output.pptx");
  const sectionBoundaryOrderAudit = path.join(sectionBoundaryDir, "order-audit.json");
  await assert.rejects(
    () => replacePptxSectionPartition({
      inputPath: sectionBoundaryInput,
      outputPath: sectionBoundaryOrderOutput,
      auditPath: sectionBoundaryOrderAudit,
      expectedSections: expectedSectionPartition,
      replacementSections: [
        { ...replacementSectionPartition[0], slideIds: ["presentation/slide/2"] },
        { ...replacementSectionPartition[1], slideIds: ["presentation/slide/1", "presentation/slide/3"] },
        replacementSectionPartition[2],
      ],
    }),
    /must partition every retained deck slide exactly once and in presentation order/i,
  );
  assert.equal(await fs.access(sectionBoundaryOrderOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(sectionBoundaryOrderAudit).then(() => true, () => false), false);
  const sectionBoundaryOpaqueInput = path.join(sectionBoundaryDir, "section-boundary-opaque-source.pptx");
  const sectionBoundaryOpaqueOutput = path.join(sectionBoundaryDir, "section-boundary-opaque-output.pptx");
  const sectionBoundaryOpaqueAudit = path.join(sectionBoundaryDir, "opaque-audit.json");
  const sectionBoundaryOpaqueZip = await JSZip.loadAsync(sectionBoundarySource);
  const sectionBoundaryPresentationXml = await sectionBoundaryOpaqueZip.file("ppt/presentation.xml").async("text");
  sectionBoundaryOpaqueZip.file("ppt/presentation.xml", sectionBoundaryPresentationXml.replace('name="Context"', 'name="Context" unmodeled="true"'));
  await fs.writeFile(sectionBoundaryOpaqueInput, await sectionBoundaryOpaqueZip.generateAsync({ type: "nodebuffer" }));
  await assert.rejects(
    () => replacePptxSectionPartition({
      inputPath: sectionBoundaryOpaqueInput,
      outputPath: sectionBoundaryOpaqueOutput,
      auditPath: sectionBoundaryOpaqueAudit,
      expectedSections: expectedSectionPartition,
      replacementSections: replacementSectionPartition,
    }),
    /no semantic canonical PowerPoint sections/i,
  );
  assert.equal(await fs.access(sectionBoundaryOpaqueOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(sectionBoundaryOpaqueAudit).then(() => true, () => false), false);
  const sectionBoundaryCollisionOutput = path.join(sectionBoundaryDir, "section-boundary-existing-output.pptx");
  const sectionBoundaryCollisionAudit = path.join(sectionBoundaryDir, "collision-audit.json");
  const sectionBoundaryCollisionSentinel = Buffer.from("preserve-existing-section-boundary-output");
  await fs.writeFile(sectionBoundaryCollisionOutput, sectionBoundaryCollisionSentinel);
  await assert.rejects(
    () => replacePptxSectionPartition({
      inputPath: sectionBoundaryInput,
      outputPath: sectionBoundaryCollisionOutput,
      auditPath: sectionBoundaryCollisionAudit,
      expectedSections: expectedSectionPartition,
      replacementSections: replacementSectionPartition,
    }),
    /outputPath already exists; refusing to overwrite/i,
  );
  assert.deepEqual(await fs.readFile(sectionBoundaryCollisionOutput), sectionBoundaryCollisionSentinel);
  assert.equal(await fs.access(sectionBoundaryCollisionAudit).then(() => true, () => false), false);
  const sectionBoundaryCliOutput = path.join(sectionBoundaryDir, "section-boundary-cli.pptx");
  const sectionBoundaryCliAudit = path.join(sectionBoundaryDir, "cli-audit.json");
  const sectionBoundaryExpectedJson = path.join(sectionBoundaryDir, "expected-sections.json");
  const sectionBoundaryReplacementJson = path.join(sectionBoundaryDir, "replacement-sections.json");
  await fs.writeFile(sectionBoundaryExpectedJson, JSON.stringify(expectedSectionPartition));
  await fs.writeFile(sectionBoundaryReplacementJson, JSON.stringify(replacementSectionPartition));
  const sectionBoundaryCli = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/examples/officekit-section-boundary-edit-workflow.mjs",
    sectionBoundaryInput,
    sectionBoundaryCliOutput,
    sectionBoundaryCliAudit,
    "@" + sectionBoundaryExpectedJson,
    "@" + sectionBoundaryReplacementJson,
  ], { encoding: "utf8" });
  assert.equal(sectionBoundaryCli.status, 0, `section-boundary CLI failed\n${sectionBoundaryCli.stdout}\n${sectionBoundaryCli.stderr}`);
  const sectionBoundaryCliResult = JSON.parse(sectionBoundaryCli.stdout);
  assert.equal(sectionBoundaryCliResult.targetPart, "ppt/presentation.xml");
  assert.equal(sectionBoundaryCliResult.changedSections, 2);
  const sectionBoundaryInlineOutput = path.join(sectionBoundaryDir, "section-boundary-inline-cli.pptx");
  const sectionBoundaryInlineAudit = path.join(sectionBoundaryDir, "inline-cli-audit.json");
  const sectionBoundaryInlineCli = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/examples/officekit-section-boundary-edit-workflow.mjs",
    sectionBoundaryInput,
    sectionBoundaryInlineOutput,
    sectionBoundaryInlineAudit,
    JSON.stringify(expectedSectionPartition),
    JSON.stringify(replacementSectionPartition),
  ], { encoding: "utf8" });
  assert.equal(sectionBoundaryInlineCli.status, 0, `section-boundary inline CLI failed\n${sectionBoundaryInlineCli.stdout}\n${sectionBoundaryInlineCli.stderr}`);
  assert.equal(JSON.parse(sectionBoundaryInlineCli.stdout).changedSections, 2);
  const sectionBoundaryOversizedExpectedJson = path.join(sectionBoundaryDir, "oversized-expected-sections.json");
  const sectionBoundaryOversizedOutput = path.join(sectionBoundaryDir, "section-boundary-oversized-cli.pptx");
  const sectionBoundaryOversizedAudit = path.join(sectionBoundaryDir, "oversized-cli-audit.json");
  await fs.writeFile(sectionBoundaryOversizedExpectedJson, "[");
  await fs.truncate(sectionBoundaryOversizedExpectedJson, (32 * 1024 * 1024) + 1);
  const sectionBoundaryOversizedCli = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/examples/officekit-section-boundary-edit-workflow.mjs",
    sectionBoundaryInput,
    sectionBoundaryOversizedOutput,
    sectionBoundaryOversizedAudit,
    "@" + sectionBoundaryOversizedExpectedJson,
    "@" + sectionBoundaryReplacementJson,
  ], { encoding: "utf8" });
  assert.notEqual(sectionBoundaryOversizedCli.status, 0, "section-boundary CLI must refuse an oversized partition file before parsing it");
  assert.match(sectionBoundaryOversizedCli.stderr, /exceeds the 33554432-byte JSON budget/i);
  assert.equal(await fs.access(sectionBoundaryOversizedOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(sectionBoundaryOversizedAudit).then(() => true, () => false), false);

  const customShowDir = path.join(root, "custom-show-workflow");
  const customShowInput = path.join(customShowDir, "custom-show-source.pptx");
  const customShowOutput = path.join(customShowDir, "custom-show-updated.pptx");
  const customShowAudit = path.join(customShowDir, "audit.json");
  await fs.mkdir(customShowDir, { recursive: true });
  const customShowFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
  const customShowSlides = [
    ["Overview", "#DBEAFE"],
    ["Evidence", "#DCFCE7"],
    ["Appendix", "#FEF3C7"],
  ].map(([name, fill], index) => {
    const slide = customShowFixture.slides.add({ name, background: { fill } });
    slide.shapes.add({
      name: `${name.toLowerCase()}-title`,
      position: { left: 44, top: 44, width: 552, height: 72 },
      text: index === 0
        ? [{ runs: [{ text: `${index + 1}. ${name}`, link: { customShow: "Board route", returnToSlide: true } }] }]
        : `${index + 1}. ${name}`,
      fill: "#FFFFFF",
      line: { fill: "#0F172A", width: 1 },
    });
    return slide;
  });
  customShowFixture.customShows.add({ name: "Board route", nativeId: 7, slides: [customShowSlides[0], customShowSlides[2]] });
  customShowFixture.customShows.add({ name: "Review route", nativeId: 11, slides: [customShowSlides[1]] });
  await (await PresentationFile.exportPptx(customShowFixture)).save(customShowInput);
  const customShowSource = await fs.readFile(customShowInput);
  const { editPptxCustomShow } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-custom-show-workflow.mjs"
  );
  const customShowResult = await editPptxCustomShow({
    inputPath: customShowInput,
    outputPath: customShowOutput,
    auditPath: customShowAudit,
    expectedName: "Board route",
    replacementName: "Executive route",
    orderedSlideNames: ["Appendix", "Overview", "Appendix"],
  });
  assert.equal(customShowResult.audit.provider.actual, "office-kit");
  assert.equal(customShowResult.audit.operation.nativeId, 7);
  assert.equal(customShowResult.audit.validation.package.onlyPresentationPartChanged, true);
  assert.equal(customShowResult.audit.validation.package.nativeIdPreserved, true);
  assert.equal(customShowResult.audit.validation.package.linkedRunCount, 1);
  assert.equal(customShowResult.audit.validation.modelRender.normalizedSvgByteIdentical, true);
  assert.deepEqual(await fs.readFile(customShowInput), customShowSource);
  const customShowRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(customShowOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "custom-show-updated.pptx",
  }));
  assert.deepEqual(customShowRoundTrip.customShows.items.map((show) => [show.name, show.nativeId]), [["Executive route", 7], ["Review route", 11]]);
  assert.deepEqual(customShowRoundTrip.customShows.getItem("Executive route").slides.map((slide) => slide.name), ["Appendix", "Overview", "Appendix"]);
  assert.equal(itemByName(customShowRoundTrip.slides.items[0].shapes.items, "overview-title").text.paragraphs[0].runs[0].link.customShow, "Executive route");
  const customShowBaselineDir = path.join(customShowDir, "baselines");
  await verifyPresentationFile(customShowInput, {
    outputDir: path.join(customShowDir, "source-qa"),
    nativeRender,
    baselineDir: customShowBaselineDir,
    writeBaseline: true,
  });
  const customShowReview = await verifyPresentationFile(customShowOutput, {
    outputDir: path.join(customShowDir, "output-qa"),
    nativeRender,
    baselineDir: customShowBaselineDir,
  });
  assert.ok(customShowReview.modelRender.slides.every((slide) => slide.pixelDiff?.changed === false));
  if (nativeStatus.available) assert.ok(customShowReview.nativeRender.pages.every((page) => page.pixelDiff?.changed === false));
  const customShowCliOutput = path.join(customShowDir, "custom-show-cli.pptx");
  const customShowCliAudit = path.join(customShowDir, "cli-audit.json");
  const customShowCli = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/examples/officekit-custom-show-workflow.mjs",
    customShowInput,
    customShowCliOutput,
    customShowCliAudit,
    "Board route",
    "CLI route",
    "Evidence,Overview",
  ], { encoding: "utf8" });
  assert.equal(customShowCli.status, 0, `custom-show CLI failed\n${customShowCli.stdout}\n${customShowCli.stderr}`);
  assert.equal(JSON.parse(customShowCli.stdout).outputPath, path.resolve(customShowCliOutput));

  const duplicateDir = path.join(root, "slide-duplicate-workflow");
  const duplicateInput = path.join(duplicateDir, "connector-source.pptx");
  const duplicateOutput = path.join(duplicateDir, "connector-duplicate.pptx");
  const duplicateAudit = path.join(duplicateDir, "audit.json");
  await fs.mkdir(duplicateDir, { recursive: true });
  const duplicateFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
  const duplicateSourceSlide = duplicateFixture.slides.add({ name: "Clone connector source" });
  const duplicateGroup = duplicateSourceSlide.addGroup({
    name: "connector-cluster",
    position: { left: 48, top: 40, width: 320, height: 120 },
    childFrame: { left: 0, top: 0, width: 320, height: 120 },
  });
  const duplicateLeft = duplicateGroup.shapes.add({ name: "left", position: { left: 0, top: 20, width: 90, height: 42 }, text: "Left" });
  const duplicateRight = duplicateGroup.shapes.add({ name: "right", position: { left: 210, top: 20, width: 90, height: 42 }, text: "Right" });
  duplicateGroup.connectors.add({
    name: "join",
    from: duplicateLeft,
    to: duplicateRight,
    start: { x: 90, y: 41 },
    end: { x: 210, y: 41 },
    line: { fill: "#64748B", width: 1 },
  });
  const duplicateCanary = duplicateFixture.slides.add({ name: "Untouched canary" });
  duplicateCanary.shapes.add({
    name: "canary-copy",
    position: { left: 48, top: 40, width: 260, height: 72 },
    text: "Unchanged source slide",
  });
  duplicateSourceSlide.shapes.add({
    name: "clone-links",
    geometry: "textbox",
    position: { left: 48, top: 190, width: 480, height: 72 },
    fill: "transparent",
    line: { fill: "transparent", width: 0 },
    text: [{ runs: [
      { text: "Guide ", link: { uri: "https://example.com/clone-guide", tooltip: "Open guide" } },
      { text: "Canary ", link: { slideId: duplicateCanary.id } },
      { text: "Next ", link: { action: "nextSlide" } },
      { text: "Board route", link: { customShow: "Clone route", returnToSlide: true, tooltip: "Open clone route" } },
    ] }],
  });
  duplicateFixture.customShows.add({
    name: "Clone route",
    nativeId: 17,
    slides: [duplicateSourceSlide, duplicateCanary],
  });
  duplicateSourceSlide.charts.add("bar", {
    name: "clone-pipeline-chart",
    position: { left: 390, top: 36, width: 218, height: 132 },
    title: "Pipeline",
    categories: ["Now", "Next"],
    series: [{ name: "Value", values: [42, 57], fill: "#2563EB" }],
    axes: { category: { title: "Stage" }, value: { title: "Value", min: 0, max: 80, majorUnit: 20 } },
    legend: false,
  });
  const duplicateSourcePptx = await PresentationFile.exportPptx(duplicateFixture);
  await duplicateSourcePptx.save(duplicateInput);
  const duplicateSourceBytes = await fs.readFile(duplicateInput);
  const { duplicatePptxSlide } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-slide-duplicate-workflow.mjs"
  );
  const duplicateResult = await duplicatePptxSlide({
    inputPath: duplicateInput,
    outputPath: duplicateOutput,
    auditPath: duplicateAudit,
    expectedName: "Clone connector source",
  });
  assert.equal(duplicateResult.audit.provider.actual, "office-kit");
  assert.equal(duplicateResult.audit.operation.type, "source-bound-slide-duplicate");
  assert.equal(duplicateResult.audit.operation.sourcePart, "ppt/slides/slide1.xml");
  assert.match(duplicateResult.audit.operation.clonePart, /^ppt\/slides\/slide\d+\.xml$/);
  assert.notEqual(duplicateResult.audit.operation.clonePart, duplicateResult.audit.operation.sourcePart);
  assert.equal(duplicateResult.audit.operation.scope, "canonical-inline-leaves-with-closed-chart-leaves");
  assert.equal(duplicateResult.audit.validation.package.retainedSourcePartsByteIdentical, true);
  assert.deepEqual(duplicateResult.audit.operation.runHyperlinks, {
    relationshipCount: 2,
    actionOnlyCount: 2,
    customShowCount: 1,
    customShowActions: [{ nativeId: 17, name: "Clone route", returnToSlide: true }],
  });
  assert.deepEqual(duplicateResult.audit.operation.customShows, {
    count: 1,
    shows: [{ name: "Clone route", nativeId: 17, slideParts: ["ppt/slides/slide1.xml", "ppt/slides/slide2.xml"] }],
  });
  assert.deepEqual(duplicateResult.audit.validation.package.runHyperlinks, {
    relationshipCount: 2,
    actionOnlyCount: 2,
    customShowCount: 1,
    customShowActions: [{ nativeId: 17, name: "Clone route", returnToSlide: true }],
    exactSourceGraphRetained: true,
  });
  assert.equal(duplicateResult.audit.validation.package.customShows.exactSourceMembershipRetained, true);
  assert.equal(duplicateResult.audit.validation.reimport.customShowMembershipRetained, true);
  assert.equal(duplicateResult.audit.operation.chartParts.count, 1);
  assert.equal(duplicateResult.audit.validation.package.chartParts.count, 1);
  assert.equal(duplicateResult.audit.validation.package.chartParts.independentParts, true);
  assert.equal(duplicateResult.audit.validation.package.chartParts.allPayloadsByteIdentical, true);
  const duplicateChartAudit = duplicateResult.audit.validation.package.chartParts.parts[0];
  assert.match(duplicateChartAudit.sourcePart, /^ppt\/(?:slides\/)?charts\/chart\d+\.xml$/i);
  assert.match(duplicateChartAudit.clonePart, /^ppt\/(?:slides\/)?charts\/chart\d+\.xml$/i);
  assert.notEqual(duplicateChartAudit.sourcePart, duplicateChartAudit.clonePart);
  const duplicateCloneRelationshipPart = duplicateResult.audit.operation.clonePart.replace(
    /(^|\/)slide(\d+)\.xml$/,
    "$1_rels/slide$2.xml.rels",
  );
  assert.deepEqual(duplicateResult.audit.validation.package.newPartPaths, [
    duplicateCloneRelationshipPart,
    duplicateChartAudit.clonePart,
    duplicateResult.audit.operation.clonePart,
  ].sort());
  assert.equal(duplicateResult.audit.validation.reimport.sourceAndCloneSemanticsEqual, true);
  assert.equal(duplicateResult.audit.validation.modelRender.visualEquivalent, true);
  assert.equal(duplicateResult.audit.validation.modelRender.identityAttributesIgnored, true);
  assert.deepEqual(await fs.readFile(duplicateInput), duplicateSourceBytes);
  const duplicateRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(duplicateOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "connector-duplicate.pptx",
  }));
  assert.deepEqual(duplicateRoundTrip.slides.items.map((slide) => slide.name), [
    "Clone connector source",
    "Clone connector source",
    "Untouched canary",
  ]);
  const duplicateSourceGroup = duplicateRoundTrip.slides.getItem(0).groups.items[0];
  const duplicateCloneGroup = duplicateRoundTrip.slides.getItem(1).groups.items[0];
  const duplicateCloneConnector = duplicateCloneGroup.connectors.items[0];
  assert.notEqual(duplicateCloneGroup.id, duplicateSourceGroup.id);
  assert.equal(duplicateCloneConnector.startTargetId, duplicateCloneGroup.shapes.items[0].id);
  assert.equal(duplicateCloneConnector.endTargetId, duplicateCloneGroup.shapes.items[1].id);
  assert.notEqual(duplicateCloneConnector.startTargetId, duplicateSourceGroup.shapes.items[0].id);
  const duplicateCloneLinks = duplicateRoundTrip.slides.getItem(1).shapes.items.find((shape) => shape.name === "clone-links").text.paragraphs[0].runs;
  assert.equal(duplicateCloneLinks[0].link.uri, "https://example.com/clone-guide");
  assert.equal(duplicateCloneLinks[1].link.slideId, duplicateRoundTrip.slides.getItem(2).id);
  assert.equal(duplicateCloneLinks[2].link.action, "nextSlide");
  assert.deepEqual(duplicateCloneLinks[3].link, {
    customShow: "Clone route",
    returnToSlide: true,
    tooltip: "Open clone route",
  });
  assert.deepEqual(duplicateRoundTrip.customShows.getItem("Clone route").slideIds, [
    duplicateRoundTrip.slides.getItem(0).id,
    duplicateRoundTrip.slides.getItem(2).id,
  ]);
  assert.ok(!duplicateRoundTrip.customShows.getItem("Clone route").slideIds.includes(duplicateRoundTrip.slides.getItem(1).id));
  const duplicateSourceChart = duplicateRoundTrip.slides.getItem(0).charts.items[0];
  const duplicateCloneChart = duplicateRoundTrip.slides.getItem(1).charts.items[0];
  assert.notEqual(duplicateCloneChart.id, duplicateSourceChart.id);
  assert.equal(duplicateSourceChart.title, "Pipeline");
  assert.equal(duplicateCloneChart.title, "Pipeline");
  assert.deepEqual(duplicateCloneChart.series[0].values, [42, 57]);
  const duplicateQa = await verifyPresentationFile(duplicateOutput, {
    outputDir: path.join(duplicateDir, "render-qa"),
    nativeRender,
  });
  assert.equal(duplicateQa.verify.ok, true);
  assert.equal(duplicateQa.packageInspect.ok, true);
  assert.equal(duplicateQa.modelRender.ok, true);
  if (nativeStatus.available) {
    assert.equal(duplicateQa.nativeRender.status, "passed");
    assert.deepEqual(
      await fs.readFile(duplicateQa.nativeRender.pages[0].path),
      await fs.readFile(duplicateQa.nativeRender.pages[1].path),
      "LibreOffice/Poppler must render the source and canonical hyperlink clone identically",
    );
  }

  duplicateCloneChart.title = "Updated clone pipeline";
  duplicateCloneChart.series[0].values[1] = 63;
  const duplicateEdited = await PresentationFile.exportPptx(duplicateRoundTrip);
  const duplicateEditedRoundTrip = await PresentationFile.importPptx(duplicateEdited);
  assert.equal(duplicateEditedRoundTrip.slides.getItem(0).charts.items[0].title, "Pipeline");
  assert.equal(duplicateEditedRoundTrip.slides.getItem(0).charts.items[0].series[0].values[1], 57);
  assert.equal(duplicateEditedRoundTrip.slides.getItem(1).charts.items[0].title, "Updated clone pipeline");
  assert.equal(duplicateEditedRoundTrip.slides.getItem(1).charts.items[0].series[0].values[1], 63);

  const oleDuplicateInput = path.join(duplicateDir, "ole-workbook-source.pptx");
  const oleDuplicateOutput = path.join(duplicateDir, "ole-workbook-output.pptx");
  const oleDuplicateAudit = path.join(duplicateDir, "ole-workbook-audit.json");
  const oleDuplicateFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
  oleDuplicateFixture.slides.add({ name: "OLE clone source" }).shapes.add({
    name: "ole-clone-copy",
    position: { left: 48, top: 40, width: 300, height: 72 },
    text: "OLE clone source",
  });
  const oleDuplicateBase = await PresentationFile.exportPptx(oleDuplicateFixture);
  const oleDuplicateBaseZip = await JSZip.loadAsync(oleDuplicateBase.bytes);
  const oleDuplicateSlideXml = await oleDuplicateBaseZip.file("ppt/slides/slide1.xml").async("text");
  const oleDuplicateRelationships = await oleDuplicateBaseZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
  const oleDuplicateWorkbook = Workbook.create();
  oleDuplicateWorkbook.worksheets.add("Embedded").getRange("A1:B2").values = [["Original clone workbook", 42], ["Status", "Ready"]];
  const oleDuplicateWorkbookFile = await SpreadsheetFile.exportXlsx(oleDuplicateWorkbook);
  const oleDuplicatePreview = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", "base64");
  const oleDuplicateFrame = '<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="100" name="Embedded clone workbook"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="1143000"/><a:ext cx="3657600" cy="1828800"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj showAsIcon="1" r:id="rIdCloneWorkbook" imgW="965200" imgH="609600" progId="Excel.Sheet.12"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdCloneWorkbookPreview"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="1143000"/><a:ext cx="3657600" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>';
  const oleDuplicateSource = await PresentationFile.patchPptx(oleDuplicateBase, [
    { path: "ppt/slides/slide1.xml", xml: oleDuplicateSlideXml.replace("</p:spTree>", `${oleDuplicateFrame}</p:spTree>`) },
    { path: "ppt/slides/_rels/slide1.xml.rels", xml: oleDuplicateRelationships.replace("</Relationships>", '<Relationship Id="rIdCloneWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/skill-clone-workbook.xlsx"/><Relationship Id="rIdCloneWorkbookPreview" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/skill-clone-workbook-preview.png"/></Relationships>') },
    { path: "ppt/embeddings/skill-clone-workbook.xlsx", bytes: oleDuplicateWorkbookFile.bytes, contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
    { path: "ppt/media/skill-clone-workbook-preview.png", bytes: oleDuplicatePreview, contentType: "image/png" },
  ]);
  await oleDuplicateSource.save(oleDuplicateInput);
  const oleDuplicateSourceBytes = await fs.readFile(oleDuplicateInput);
  const oleDuplicateResult = await duplicatePptxSlide({
    inputPath: oleDuplicateInput,
    outputPath: oleDuplicateOutput,
    auditPath: oleDuplicateAudit,
    expectedName: "OLE clone source",
  });
  assert.equal(oleDuplicateResult.audit.operation.scope, "canonical-inline-leaves-with-closed-ole-workbook-leaves");
  assert.deepEqual(oleDuplicateResult.audit.operation.oleWorkbookParts, {
    count: 1,
    sourceParts: ["ppt/embeddings/skill-clone-workbook.xlsx"],
    relationshipIds: ["rIdCloneWorkbook"],
    previewParts: ["ppt/media/skill-clone-workbook-preview.png"],
  });
  const olePackageAudit = oleDuplicateResult.audit.validation.package.oleWorkbookParts;
  assert.equal(olePackageAudit.count, 1);
  assert.equal(olePackageAudit.independentParts, true);
  assert.equal(olePackageAudit.allPayloadsByteIdentical, true);
  assert.equal(olePackageAudit.previewPartsShared, true);
  const [olePartAudit] = olePackageAudit.parts;
  assert.equal(olePartAudit.sourcePart, "ppt/embeddings/skill-clone-workbook.xlsx");
  assert.match(olePartAudit.clonePart, /^ppt\/embeddings\/[^/]+\.xlsx$/i);
  assert.notEqual(olePartAudit.clonePart, olePartAudit.sourcePart);
  assert.equal(olePartAudit.relationshipId, "rIdCloneWorkbook");
  assert.equal(olePartAudit.previewRelationshipId, "rIdCloneWorkbookPreview");
  assert.equal(olePartAudit.previewPart, "ppt/media/skill-clone-workbook-preview.png");
  assert.equal(olePartAudit.workbookBytesByteIdentical, true);
  assert.equal(olePartAudit.previewShared, true);
  assert.deepEqual(oleDuplicateResult.audit.validation.package.newPartPaths, [
    relationshipPartPath(oleDuplicateResult.audit.operation.clonePart),
    olePartAudit.clonePart,
    oleDuplicateResult.audit.operation.clonePart,
  ].sort());
  assert.equal(oleDuplicateResult.audit.validation.reimport.sourceAndCloneOleWorkbookBindingsIndependent, true);
  assert.deepEqual(await fs.readFile(oleDuplicateInput), oleDuplicateSourceBytes);
  const oleDuplicateRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(oleDuplicateOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "ole-workbook-output.pptx",
  }));
  const oleSourceObject = itemByName(oleDuplicateRoundTrip.slides.getItem(0).nativeObjects.items, "Embedded clone workbook");
  const oleCloneObject = itemByName(oleDuplicateRoundTrip.slides.getItem(1).nativeObjects.items, "Embedded clone workbook");
  assert.notEqual(oleCloneObject.id, oleSourceObject.id);
  assert.notEqual(oleCloneObject.oleWorkbook.partPath, oleSourceObject.oleWorkbook.partPath);
  assert.equal(oleCloneObject.oleWorkbook.sourceSha256, oleSourceObject.oleWorkbook.sourceSha256);
  const oleReplacementWorkbook = Workbook.create();
  oleReplacementWorkbook.worksheets.add("Embedded").getRange("A1:B2").values = [["Independent clone workbook", 84], ["Status", "Edited"]];
  const oleReplacementFile = await SpreadsheetFile.exportXlsx(oleReplacementWorkbook);
  oleCloneObject.replaceEmbeddedWorkbook(oleReplacementFile);
  const oleEditedFile = await PresentationFile.exportPptx(oleDuplicateRoundTrip);
  const oleEditedZip = await JSZip.loadAsync(oleEditedFile.bytes);
  assert.deepEqual(await oleEditedZip.file(oleSourceObject.oleWorkbook.partPath).async("uint8array"), oleDuplicateWorkbookFile.bytes);
  assert.deepEqual(await oleEditedZip.file(oleCloneObject.oleWorkbook.partPath).async("uint8array"), oleReplacementFile.bytes);
  const oleEditedRoundTrip = await PresentationFile.importPptx(oleEditedFile);
  const oleEditedSourceWorkbook = await SpreadsheetFile.importXlsx(itemByName(oleEditedRoundTrip.slides.getItem(0).nativeObjects.items, "Embedded clone workbook").getEmbeddedWorkbook());
  const oleEditedCloneWorkbook = await SpreadsheetFile.importXlsx(itemByName(oleEditedRoundTrip.slides.getItem(1).nativeObjects.items, "Embedded clone workbook").getEmbeddedWorkbook());
  assert.equal(oleEditedSourceWorkbook.worksheets.getItem("Embedded").getRange("A1").values[0][0], "Original clone workbook");
  assert.equal(oleEditedCloneWorkbook.worksheets.getItem("Embedded").getRange("A1").values[0][0], "Independent clone workbook");

  const oleOfficePackageDir = path.join(root, "ole-office-package-workflow");
  const oleOfficePackageInput = path.join(oleOfficePackageDir, "source.pptx");
  const oleOfficePackageOutput = path.join(oleOfficePackageDir, "edited.pptx");
  const oleOfficePackageAudit = path.join(oleOfficePackageDir, "audit.json");
  const oleOfficePackageDocument = DocumentModel.create({ name: "Embedded DOCX source", blocks: [] });
  oleOfficePackageDocument.addParagraph("Draft approval wording");
  const oleOfficePackageDocx = await DocumentFile.exportDocx(oleOfficePackageDocument);
  const oleOfficePackageDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
  oleOfficePackageDeck.slides.add({ name: "DOCX package source" }).shapes.add({
    name: "ole-package-copy",
    position: { left: 48, top: 40, width: 300, height: 72 },
    text: "DOCX package source",
  });
  const oleOfficePackageBase = await PresentationFile.exportPptx(oleOfficePackageDeck);
  const oleOfficePackageBaseZip = await JSZip.loadAsync(oleOfficePackageBase.bytes);
  const oleOfficePackageSlideXml = await oleOfficePackageBaseZip.file("ppt/slides/slide1.xml").async("text");
  const oleOfficePackageRelationships = await oleOfficePackageBaseZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
  const oleOfficePackageFrame = '<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="170" name="Embedded approval document"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="1143000"/><a:ext cx="3657600" cy="1828800"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj showAsIcon="1" r:id="rIdSkillEmbeddedDocument" imgW="965200" imgH="609600" progId="Word.Document.12"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdSkillEmbeddedDocumentPreview"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="1143000"/><a:ext cx="3657600" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>';
  const oleOfficePackagePreview = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", "base64");
  const oleOfficePackageSource = await PresentationFile.patchPptx(oleOfficePackageBase, [
    { path: "ppt/slides/slide1.xml", xml: oleOfficePackageSlideXml.replace("</p:spTree>", `${oleOfficePackageFrame}</p:spTree>`) },
    { path: "ppt/slides/_rels/slide1.xml.rels", xml: oleOfficePackageRelationships.replace("</Relationships>", '<Relationship Id="rIdSkillEmbeddedDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/skill-approval.docx"/><Relationship Id="rIdSkillEmbeddedDocumentPreview" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/skill-approval-preview.png"/></Relationships>') },
    { path: "ppt/embeddings/skill-approval.docx", bytes: oleOfficePackageDocx.bytes, contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
    { path: "ppt/media/skill-approval-preview.png", bytes: oleOfficePackagePreview, contentType: "image/png" },
  ]);
  await fs.mkdir(oleOfficePackageDir, { recursive: true });
  await oleOfficePackageSource.save(oleOfficePackageInput);
  const oleOfficePackageSourceBytes = await fs.readFile(oleOfficePackageInput);
  const { editPptxEmbeddedDocxPackage } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-ole-office-package-workflow.mjs"
  );
  const oleOfficePackageResult = await editPptxEmbeddedDocxPackage({
    inputPath: oleOfficePackageInput,
    outputPath: oleOfficePackageOutput,
    auditPath: oleOfficePackageAudit,
    objectName: "Embedded approval document",
    expectedText: "Draft approval wording",
    replacementText: "Approved approval wording",
  });
  assert.equal(oleOfficePackageResult.audit.provider.actual, "office-kit");
  assert.equal(oleOfficePackageResult.audit.provider.silentFallback, false);
  assert.equal(oleOfficePackageResult.audit.operation.type, "source-bound-ole-docx-package-paragraph-edit");
  assert.equal(oleOfficePackageResult.audit.operation.package.kind, "docx");
  assert.equal(oleOfficePackageResult.audit.operation.package.partPath, "ppt/embeddings/skill-approval.docx");
  assert.deepEqual(oleOfficePackageResult.audit.validation.package.changedPartPaths, ["ppt/embeddings/skill-approval.docx"]);
  assert.equal(oleOfficePackageResult.audit.validation.package.nonTargetPartsByteIdentical, true);
  assert.equal(oleOfficePackageResult.audit.validation.package.exactReplacementBytes, true);
  assert.equal(oleOfficePackageResult.audit.validation.reimport.sourceBindingReproved, true);
  assert.equal(oleOfficePackageResult.audit.validation.modelRender.byteIdentical, true);
  assert.deepEqual(await fs.readFile(oleOfficePackageInput), oleOfficePackageSourceBytes);
  const oleOfficePackageRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(oleOfficePackageOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "edited.pptx",
  }));
  const reboundOfficePackageObject = itemByName(oleOfficePackageRoundTrip.slides.getItem(0).nativeObjects.items, "Embedded approval document");
  assert.equal(reboundOfficePackageObject.oleOfficePackage.kind, "docx");
  assert.equal(reboundOfficePackageObject.inspectRecord().embeddedOfficePackage.replacementPending, false);
  assert.deepEqual((await DocumentFile.importDocx(reboundOfficePackageObject.getEmbeddedOfficePackage())).paragraphs, ["Approved approval wording"]);
  const oleOfficePackageOutputZip = await JSZip.loadAsync(await fs.readFile(oleOfficePackageOutput));
  assert.deepEqual(await oleOfficePackageOutputZip.file("ppt/media/skill-approval-preview.png").async("uint8array"), Uint8Array.from(oleOfficePackagePreview));

  const smartArtDuplicateInput = path.join(duplicateDir, "smartart-source.pptx");
  const smartArtDuplicateOutput = path.join(duplicateDir, "smartart-output.pptx");
  const smartArtDuplicateAudit = path.join(duplicateDir, "smartart-audit.json");
  const smartArtFrame = '<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="120" name="Closed SmartArt"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="1143000"/><a:ext cx="4572000" cy="2286000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdSkillDiagramData" r:lo="rIdSkillDiagramLayout" r:qs="rIdSkillDiagramStyle" r:cs="rIdSkillDiagramColors"/></a:graphicData></a:graphic></p:graphicFrame>';
  const smartArtRelationships = '<Relationship Id="rIdSkillDiagramData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/skill-data.xml"/><Relationship Id="rIdSkillDiagramLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/skill-layout.xml"/><Relationship Id="rIdSkillDiagramStyle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/skill-style.xml"/><Relationship Id="rIdSkillDiagramColors" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/skill-colors.xml"/>';
  const smartArtDataXml = '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="{B59B8E5A-4DF0-4A3C-A5E2-A7D7B293E601}" type="doc"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Original SmartArt node</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="{C6D16D59-0A5A-42E6-AF7F-C53A0E3D487C}" type="doc"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Second SmartArt node</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst/><dgm:bg/><dgm:whole/></dgm:dataModel>';
  const smartArtSource = await PresentationFile.patchPptx(oleDuplicateBase, [
    { path: "ppt/slides/slide1.xml", xml: oleDuplicateSlideXml.replace('name="OLE clone source"', 'name="SmartArt clone source"').replace("</p:spTree>", `${smartArtFrame}</p:spTree>`) },
    { path: "ppt/slides/_rels/slide1.xml.rels", xml: oleDuplicateRelationships.replace("</Relationships>", `${smartArtRelationships}</Relationships>`) },
    { path: "ppt/diagrams/skill-data.xml", contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml", xml: smartArtDataXml },
    { path: "ppt/diagrams/skill-layout.xml", contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml", xml: '<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:office-kit:skill-layout"><dgm:title val="Skill"/><dgm:desc val="Skill layout"/><dgm:catLst/><dgm:layoutNode name="root"/></dgm:layoutDef>' },
    { path: "ppt/diagrams/skill-style.xml", contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml", xml: '<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:office-kit:skill-style"><dgm:title val="Skill"/><dgm:desc val="Skill style"/><dgm:catLst/><dgm:styleLbl name="node0"/></dgm:styleDef>' },
    { path: "ppt/diagrams/skill-colors.xml", contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml", xml: '<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:office-kit:skill-colors"><dgm:title val="Skill"/><dgm:desc val="Skill colors"/><dgm:catLst/></dgm:colorsDef>' },
  ]);
  await smartArtSource.save(smartArtDuplicateInput);
  const smartArtSourceBytes = await fs.readFile(smartArtDuplicateInput);
  const smartArtTextInput = path.join(duplicateDir, "smartart-text-source.pptx");
  const smartArtTextOutput = path.join(duplicateDir, "smartart-text-output.pptx");
  const smartArtTextAudit = path.join(duplicateDir, "smartart-text-audit.json");
  await smartArtSource.save(smartArtTextInput);
  const { editPptxSmartArtNodeText } = await import(
    "../skills/presentations/skills/presentations/examples/officekit-smartart-text-edit-workflow.mjs"
  );
  const smartArtTextResult = await editPptxSmartArtNodeText({
    inputPath: smartArtTextInput,
    outputPath: smartArtTextOutput,
    auditPath: smartArtTextAudit,
    objectName: "Closed SmartArt",
    nodeId: "{B59B8E5A-4DF0-4A3C-A5E2-A7D7B293E601}",
    expectedText: "Original SmartArt node",
    replacementText: "Updated SmartArt node",
  });
  assert.equal(smartArtTextResult.audit.provider.actual, "office-kit");
  assert.equal(smartArtTextResult.audit.operation.type, "source-bound-smartart-node-text-edit");
  assert.equal(smartArtTextResult.audit.operation.dataPart, "ppt/diagrams/skill-data.xml");
  assert.deepEqual(smartArtTextResult.audit.validation.package.changedPartPaths, ["ppt/diagrams/skill-data.xml"]);
  assert.equal(smartArtTextResult.audit.validation.package.nonTargetPartsByteIdentical, true);
  assert.deepEqual(await fs.readFile(smartArtTextInput), smartArtSourceBytes);
  const smartArtTextRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(smartArtTextOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "smartart-text-output.pptx",
  }));
  assert.deepEqual(itemByName(smartArtTextRoundTrip.slides.getItem(0).nativeObjects.items, "Closed SmartArt").diagramText.nodes, [
    { id: "{B59B8E5A-4DF0-4A3C-A5E2-A7D7B293E601}", text: "Updated SmartArt node", runs: ["Updated SmartArt node"] },
    { id: "{C6D16D59-0A5A-42E6-AF7F-C53A0E3D487C}", text: "Second SmartArt node", runs: ["Second SmartArt node"] },
  ]);
  const smartArtRunInput = path.join(duplicateDir, "smartart-run-source.pptx");
  const smartArtRunOutput = path.join(duplicateDir, "smartart-run-output.pptx");
  const smartArtRunAudit = path.join(duplicateDir, "smartart-run-audit.json");
  const smartArtRunSource = await PresentationFile.patchPptx(smartArtSource, [{
    path: "ppt/diagrams/skill-data.xml",
    xml: smartArtDataXml.replace(
      "<a:r><a:t>Original SmartArt node</a:t></a:r>",
      '<a:pPr algn="ctr"/><a:r><a:rPr b="1"/><a:t>Original</a:t></a:r><a:br><a:rPr lang="fr-FR"/></a:br><a:endParaRPr lang="en-US"/></a:p><a:p><a:pPr marL="91440"/><a:endParaRPr lang="de-DE"/></a:p><a:p><a:r><a:rPr i="1"/><a:t> SmartArt node</a:t></a:r>',
    ),
  }]);
  await smartArtRunSource.save(smartArtRunInput);
  const smartArtRunResult = await editPptxSmartArtNodeText({
    inputPath: smartArtRunInput,
    outputPath: smartArtRunOutput,
    auditPath: smartArtRunAudit,
    objectName: "Closed SmartArt",
    nodeId: "{B59B8E5A-4DF0-4A3C-A5E2-A7D7B293E601}",
    runIndex: 0,
    expectedText: "Original",
    replacementText: "Updated",
  });
  assert.equal(smartArtRunResult.audit.operation.type, "source-bound-smartart-run-text-edit");
  assert.equal(smartArtRunResult.audit.operation.runIndex, 0);
  assert.equal(smartArtRunResult.audit.validation.reimport.runTopologyPreserved, true);
  const smartArtRunRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(smartArtRunOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "smartart-run-output.pptx",
  }));
  assert.deepEqual(
    itemByName(smartArtRunRoundTrip.slides.getItem(0).nativeObjects.items, "Closed SmartArt").diagramText.nodes[0],
    { id: "{B59B8E5A-4DF0-4A3C-A5E2-A7D7B293E601}", text: "Updated SmartArt node", runs: ["Updated", " SmartArt node"] },
  );
  const smartArtRunOutputZip = await JSZip.loadAsync(await fs.readFile(smartArtRunOutput));
  const smartArtRunOutputXml = await smartArtRunOutputZip.file("ppt/diagrams/skill-data.xml").async("text");
  assert.equal((smartArtRunOutputXml.match(/<a:p(?:\s|>)/g) || []).length, 4);
  assert.match(smartArtRunOutputXml, /<a:pPr algn="ctr"\s*\/>/);
  assert.match(smartArtRunOutputXml, /<a:p><a:pPr marL="91440"\s*\/><a:endParaRPr lang="de-DE"\s*\/><\/a:p>/);
  assert.match(smartArtRunOutputXml, /<a:br><a:rPr lang="fr-FR"\s*\/><\/a:br>/);
  assert.match(smartArtRunOutputXml, /<a:endParaRPr lang="en-US"\s*\/>/);
  await assert.rejects(
    () => editPptxSmartArtNodeText({
      inputPath: smartArtRunInput,
      outputPath: path.join(duplicateDir, "smartart-null-run-rejected.pptx"),
      auditPath: path.join(duplicateDir, "smartart-null-run-rejected.json"),
      objectName: "Closed SmartArt",
      nodeId: "{B59B8E5A-4DF0-4A3C-A5E2-A7D7B293E601}",
      runIndex: null,
      expectedText: "Original",
      replacementText: "No output",
    }),
    /runIndex must be a non-negative integer/,
  );
  await assert.rejects(
    () => editPptxSmartArtNodeText({
      inputPath: smartArtTextInput,
      outputPath: path.join(duplicateDir, "smartart-text-rejected.pptx"),
      auditPath: path.join(duplicateDir, "smartart-text-rejected.json"),
      objectName: "Closed SmartArt",
      nodeId: "missing-node",
      expectedText: "Original SmartArt node",
      replacementText: "No output",
    }),
    /exactly one source-bound SmartArt/,
  );
  const smartArtDuplicateResult = await duplicatePptxSlide({
    inputPath: smartArtDuplicateInput,
    outputPath: smartArtDuplicateOutput,
    auditPath: smartArtDuplicateAudit,
    expectedName: "SmartArt clone source",
  });
  assert.equal(smartArtDuplicateResult.audit.operation.scope, "canonical-inline-leaves-with-closed-smartart-leaves");
  assert.deepEqual(smartArtDuplicateResult.audit.operation.diagramParts, {
    count: 1,
    partCount: 4,
    sourceParts: [
      "ppt/diagrams/skill-data.xml",
      "ppt/diagrams/skill-layout.xml",
      "ppt/diagrams/skill-style.xml",
      "ppt/diagrams/skill-colors.xml",
    ],
    relationshipIds: [
      "rIdSkillDiagramData",
      "rIdSkillDiagramLayout",
      "rIdSkillDiagramStyle",
      "rIdSkillDiagramColors",
    ],
  });
  const smartArtPackageAudit = smartArtDuplicateResult.audit.validation.package.diagramParts;
  assert.equal(smartArtPackageAudit.count, 1);
  assert.equal(smartArtPackageAudit.partCount, 4);
  assert.equal(smartArtPackageAudit.independentParts, true);
  assert.equal(smartArtPackageAudit.allPayloadsByteIdentical, true);
  assert.equal(smartArtPackageAudit.parts.length, 4);
  assert.ok(smartArtPackageAudit.parts.every((part) => part.independentPart && part.diagramXmlByteIdentical && part.sourcePart !== part.clonePart));
  assert.ok(smartArtPackageAudit.parts.every((part) => /^ppt\/diagrams\/[^/]+\.xml$/i.test(part.clonePart)));
  assert.deepEqual(smartArtDuplicateResult.audit.validation.package.newPartPaths, [
    relationshipPartPath(smartArtDuplicateResult.audit.operation.clonePart),
    ...smartArtPackageAudit.parts.map((part) => part.clonePart),
    smartArtDuplicateResult.audit.operation.clonePart,
  ].sort());
  assert.equal(smartArtDuplicateResult.audit.validation.reimport.sourceAndCloneDiagramBindingsIndependent, true);
  assert.deepEqual(await fs.readFile(smartArtDuplicateInput), smartArtSourceBytes);
  const smartArtRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(smartArtDuplicateOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "smartart-output.pptx",
  }));
  const smartArtSourceObject = itemByName(smartArtRoundTrip.slides.getItem(0).nativeObjects.items, "Closed SmartArt");
  const smartArtCloneObject = itemByName(smartArtRoundTrip.slides.getItem(1).nativeObjects.items, "Closed SmartArt");
  assert.equal(smartArtSourceObject.parts.length, 4);
  assert.equal(smartArtCloneObject.parts.length, 4);
  assert.equal(smartArtSourceObject.parts.some((part) => smartArtCloneObject.parts.some((clonePart) => clonePart.path === part.path)), false);
  assert.deepEqual(smartArtSourceObject.parts.map((part) => part.sourceSha256).sort(), smartArtCloneObject.parts.map((part) => part.sourceSha256).sort());
  const smartArtQa = await verifyPresentationFile(smartArtDuplicateOutput, {
    outputDir: path.join(duplicateDir, "smartart-render-qa"),
    nativeRender,
  });
  assert.equal(smartArtQa.verify.ok, true);
  assert.equal(smartArtQa.packageInspect.ok, true);
  assert.equal(smartArtQa.modelRender.ok, true);
  if (nativeStatus.available) {
    assert.equal(smartArtQa.nativeRender.status, "passed");
    assert.deepEqual(
      await fs.readFile(smartArtQa.nativeRender.pages[0].path),
      await fs.readFile(smartArtQa.nativeRender.pages[1].path),
      "LibreOffice/Poppler must render the source and closed SmartArt clone identically",
    );
  }

  const connectedSmartArtInput = path.join(duplicateDir, "connected-smartart-source.pptx");
  const connectedSmartArtOutput = path.join(duplicateDir, "connected-smartart-output.pptx");
  const connectedSmartArtAudit = path.join(duplicateDir, "connected-smartart-audit.json");
  const connectedSmartArtZip = await JSZip.loadAsync(smartArtSourceBytes);
  connectedSmartArtZip.file(
    "ppt/diagrams/_rels/skill-data.xml.rels",
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdUnsafeSkillDiagram" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/smartart" TargetMode="External"/></Relationships>',
  );
  await fs.writeFile(connectedSmartArtInput, await connectedSmartArtZip.generateAsync({ type: "nodebuffer" }));
  await assert.rejects(
    () => duplicatePptxSlide({
      inputPath: connectedSmartArtInput,
      outputPath: connectedSmartArtOutput,
      auditPath: connectedSmartArtAudit,
      expectedName: "SmartArt clone source",
    }),
    /SmartArt diagramData leaf must not have a child relationship graph/,
  );
  assert.equal(await fs.access(connectedSmartArtOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(connectedSmartArtAudit).then(() => true, () => false), false);

  const inkDuplicateInput = path.join(duplicateDir, "inkml-source.pptx");
  const inkDuplicateOutput = path.join(duplicateDir, "inkml-output.pptx");
  const inkDuplicateAudit = path.join(duplicateDir, "inkml-audit.json");
  const inkContentPart = '<p:contentPart xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdSkillInk"><p14:nvContentPartPr><p14:cNvPr id="121" name="Closed InkML"/><p14:cNvContentPartPr/><p14:nvPr/></p14:nvContentPartPr><p14:xfrm><a:off x="914400" y="1143000"/><a:ext cx="4572000" cy="2286000"/></p14:xfrm></p:contentPart>';
  const inkRelationship = '<Relationship Id="rIdSkillInk" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/skill-ink.xml"/>';
  const inkSource = await PresentationFile.patchPptx(oleDuplicateBase, [
    { path: "ppt/slides/slide1.xml", xml: oleDuplicateSlideXml.replace('name="OLE clone source"', 'name="InkML clone source"').replace("</p:spTree>", `${inkContentPart}</p:spTree>`) },
    { path: "ppt/slides/_rels/slide1.xml.rels", xml: oleDuplicateRelationships.replace("</Relationships>", `${inkRelationship}</Relationships>`) },
    { path: "ppt/customXml/skill-ink.xml", contentType: "application/inkml+xml", xml: '<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:trace>0 0, 100 100, 200 0</inkml:trace></inkml:ink>' },
  ]);
  await inkSource.save(inkDuplicateInput);
  const inkSourceBytes = await fs.readFile(inkDuplicateInput);
  const inkDuplicateResult = await duplicatePptxSlide({
    inputPath: inkDuplicateInput,
    outputPath: inkDuplicateOutput,
    auditPath: inkDuplicateAudit,
    expectedName: "InkML clone source",
  });
  assert.equal(inkDuplicateResult.audit.operation.scope, "canonical-inline-leaves-with-closed-inkml-leaves");
  assert.deepEqual(inkDuplicateResult.audit.operation.inkContentParts, {
    count: 1,
    sourceParts: ["ppt/customXml/skill-ink.xml"],
    relationshipIds: ["rIdSkillInk"],
  });
  const inkPackageAudit = inkDuplicateResult.audit.validation.package.inkContentParts;
  assert.equal(inkPackageAudit.count, 1);
  assert.equal(inkPackageAudit.independentParts, true);
  assert.equal(inkPackageAudit.allPayloadsByteIdentical, true);
  const [inkPartAudit] = inkPackageAudit.parts;
  assert.equal(inkPartAudit.sourcePart, "ppt/customXml/skill-ink.xml");
  assert.match(inkPartAudit.clonePart, /^ppt\/customXml\/[^/]+\.xml$/i);
  assert.equal(inkPartAudit.relationshipId, "rIdSkillInk");
  assert.equal(inkPartAudit.inkXmlByteIdentical, true);
  assert.equal(inkPartAudit.independentPart, true);
  assert.deepEqual(inkDuplicateResult.audit.validation.package.newPartPaths, [
    relationshipPartPath(inkDuplicateResult.audit.operation.clonePart),
    inkPartAudit.clonePart,
    inkDuplicateResult.audit.operation.clonePart,
  ].sort());
  assert.equal(inkDuplicateResult.audit.validation.reimport.sourceAndCloneInkContentBindingsIndependent, true);
  assert.deepEqual(await fs.readFile(inkDuplicateInput), inkSourceBytes);
  const inkRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(inkDuplicateOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "inkml-output.pptx",
  }));
  const inkSourceObject = itemByName(inkRoundTrip.slides.getItem(0).nativeObjects.items, "Closed InkML");
  const inkCloneObject = itemByName(inkRoundTrip.slides.getItem(1).nativeObjects.items, "Closed InkML");
  assert.equal(inkSourceObject.parts.length, 1);
  assert.equal(inkCloneObject.parts.length, 1);
  assert.notEqual(inkSourceObject.parts[0].path, inkCloneObject.parts[0].path);
  assert.equal(inkSourceObject.parts[0].sourceSha256, inkCloneObject.parts[0].sourceSha256);
  const inkQa = await verifyPresentationFile(inkDuplicateOutput, {
    outputDir: path.join(duplicateDir, "inkml-render-qa"),
    nativeRender,
  });
  assert.equal(inkQa.verify.ok, true);
  assert.equal(inkQa.packageInspect.ok, true);
  assert.equal(inkQa.modelRender.ok, true);
  if (nativeStatus.available) {
    assert.equal(inkQa.nativeRender.status, "passed");
    assert.deepEqual(
      await fs.readFile(inkQa.nativeRender.pages[0].path),
      await fs.readFile(inkQa.nativeRender.pages[1].path),
      "LibreOffice/Poppler must render the source and closed InkML clone identically",
    );
  }

  const mediaDuplicateInput = path.join(duplicateDir, "embedded-video-source.pptx");
  const mediaDuplicateOutput = path.join(duplicateDir, "embedded-video-output.pptx");
  const mediaDuplicateAudit = path.join(duplicateDir, "embedded-video-audit.json");
  const mediaPayload = Buffer.from("AAAAIGZ0eXBpc29tAAACAGlzb21pc28yYXZjMW1wNDEAAAMVbW9vdgAAAGxtdmhkAAAAAAAAAAAAAAAAAAAD6AAAACgAAQAAAQAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAj90cmFrAAAAXHRraGQAAAADAAAAAAAAAAAAAAABAAAAAAAAACgAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAABAAAAAABAAAAAQAAAAAAAkZWR0cwAAABxlbHN0AAAAAAAAAAEAAAAoAAAAAAABAAAAAAG3bWRpYQAAACBtZGhkAAAAAAAAAAAAAAAAAAAyAAAAAgBVxAAAAAAALWhkbHIAAAAAAAAAAHZpZGUAAAAAAAAAAAAAAABWaWRlb0hhbmRsZXIAAAABYm1pbmYAAAAUdm1oZAAAAAEAAAAAAAAAAAAAACRkaW5mAAAAHGRyZWYAAAAAAAAAAQAAAAx1cmwgAAAAAQAAASJzdGJsAAAAvnN0c2QAAAAAAAAAAQAAAK5hdmMxAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAABAAEABIAAAASAAAAAAAAAABFUxhdmM2Mi4yOC4xMDIgbGlieDI2NAAAAAAAAAAAAAAAGP//AAAANGF2Y0MBZAAK/+EAF2dkAAqs2V7ARAAAAwAEAAADAMg8SJZYAQAGaOvjyyLA/fj4AAAAABBwYXNwAAAAAQAAAAEAAAAUYnRydAAAAAAAAinoAAAAAAAAABhzdHRzAAAAAAAAAAEAAAABAAACAAAAABxzdHNjAAAAAAAAAAEAAAABAAAAAQAAAAEAAAAUc3RzegAAAAAAAALFAAAAAQAAABRzdGNvAAAAAAAAAAEAAANFAAAAYnVkdGEAAABabWV0YQAAAAAAAAAhaGRscgAAAAAAAAAAbWRpcmFwcGwAAAAAAAAAAAAAAAAtaWxzdAAAACWpdG9vAAAAHWRhdGEAAAABAAAAAExhdmY2Mi4xMi4xMDIAAAAIZnJlZQAAAs1tZGF0AAACrgYF//+q3EXpvebZSLeWLNgg2SPu73gyNjQgLSBjb3JlIDE2NSByMzIyMiBiMzU2MDVhIC0gSC4yNjQvTVBFRy00IEFWQyBjb2RlYyAtIENvcHlsZWZ0IDIwMDMtMjAyNSAtIGh0dHA6Ly93d3cudmlkZW9sYW4ub3JnL3gyNjQuaHRtbCAtIG9wdGlvbnM6IGNhYmFjPTEgcmVmPTMgZGVibG9jaz0xOjA6MCBhbmFseXNpPTB4MzoweDExMyBtZT1oZXggc3VibWU9NyBwc3k9MSBwc3lfcmQ9MS4wMDowLjAwIG1peGVkX3JlZj0xIG1lX3JhbmdlPTE2IGNocm9tYV9tZT0xIHRyZWxsaXM9MSA4eDhkY3Q9MSBjcW09MCBkZWFkem9uZT0yMSwxMSBmYXN0X3Bza2lwPTEgY2hyb21hX3FwX29mZnNldD0tMiB0aHJlYWRzPTEgbG9va2FoZWFkX3RocmVhZHM9MSBzbGljZWRfdGhyZWFkcz0wIG5yPTAgZGVjaW1hdGU9MSBpbnRlcmxhY2VkPTAgYmx1cmF5X2NvbXBhdD0wIGNvbnN0cmFpbmVkX2ludHJhPTAgYmZyYW1lcz0zIGJfcHlyYW1pZD0yIGJfYWRhcHQ9MSBiX2JpYXM9MCBkaXJlY3Q9MSB3ZWlnaHRiPTEgb3Blbl9nb3A9MCB3ZWlnaHRwPTIga2V5aW50PTI1MCBrZXlpbnRfbWluPTI1IHNjZW5lY3V0PTQwIGludHJhX3JlZnJlc2g9MCByY19sb29rYWhlYWQ9NDAgcmM9Y3JmIG1idHJlZT0xIGNyZj0yMy4wIHFjb21wPTAuNjAgcXBtaW49MCBxcG1heD02OSBxcHN0ZXA9NCBpcF9yYXRpbz0xLjQwIGFxPTE6MS4wMACAAAAAD2WIhAAr//72c3wKa22xgQ==", "base64");
  const mediaPicture = '<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvPicPr><p:cNvPr id="122" name="Closed embedded video"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:videoFile r:link="rIdSkillVideo"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media r:embed="rIdSkillMedia"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="rIdSkillVideoPoster"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="1143000"/><a:ext cx="3657600" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
  const mediaRelationships = '<Relationship Id="rIdSkillVideo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/skill-video.mp4"/><Relationship Id="rIdSkillMedia" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/skill-video.mp4"/><Relationship Id="rIdSkillVideoPoster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/skill-video-poster.png"/>';
  const mediaSource = await PresentationFile.patchPptx(oleDuplicateBase, [
    { path: "ppt/slides/slide1.xml", xml: oleDuplicateSlideXml.replace('name="OLE clone source"', 'name="Embedded video clone source"').replace("</p:spTree>", `${mediaPicture}</p:spTree>`) },
    { path: "ppt/slides/_rels/slide1.xml.rels", xml: oleDuplicateRelationships.replace("</Relationships>", `${mediaRelationships}</Relationships>`) },
    { path: "ppt/media/skill-video.mp4", bytes: mediaPayload, contentType: "video/mp4" },
    { path: "ppt/media/skill-video-poster.png", bytes: oleDuplicatePreview, contentType: "image/png" },
  ]);
  await mediaSource.save(mediaDuplicateInput);
  const mediaSourceBytes = await fs.readFile(mediaDuplicateInput);
  const mediaDuplicateResult = await duplicatePptxSlide({
    inputPath: mediaDuplicateInput,
    outputPath: mediaDuplicateOutput,
    auditPath: mediaDuplicateAudit,
    expectedName: "Embedded video clone source",
  });
  assert.equal(mediaDuplicateResult.audit.operation.scope, "canonical-inline-leaves-with-closed-mp4-leaves");
  assert.deepEqual(mediaDuplicateResult.audit.operation.mediaParts, {
    count: 1,
    sourceParts: ["ppt/media/skill-video.mp4"],
    videoRelationshipIds: ["rIdSkillVideo"],
    mediaRelationshipIds: ["rIdSkillMedia"],
    posterParts: ["ppt/media/skill-video-poster.png"],
  });
  const mediaPackageAudit = mediaDuplicateResult.audit.validation.package.mediaParts;
  assert.equal(mediaPackageAudit.count, 1);
  assert.equal(mediaPackageAudit.independentParts, true);
  assert.equal(mediaPackageAudit.allPayloadsByteIdentical, true);
  assert.equal(mediaPackageAudit.posterPartsShared, true);
  const [mediaPartAudit] = mediaPackageAudit.parts;
  assert.equal(mediaPartAudit.sourcePart, "ppt/media/skill-video.mp4");
  assert.match(mediaPartAudit.clonePart, /^(?:ppt\/)?media\/[^/]+\.mp4$/i);
  assert.notEqual(mediaPartAudit.clonePart, mediaPartAudit.sourcePart);
  assert.equal(mediaPartAudit.videoRelationshipId, "rIdSkillVideo");
  assert.equal(mediaPartAudit.mediaRelationshipId, "rIdSkillMedia");
  assert.equal(mediaPartAudit.posterRelationshipId, "rIdSkillVideoPoster");
  assert.equal(mediaPartAudit.posterPart, "ppt/media/skill-video-poster.png");
  assert.equal(mediaPartAudit.mediaBytesByteIdentical, true);
  assert.equal(mediaPartAudit.posterShared, true);
  assert.deepEqual(mediaDuplicateResult.audit.validation.package.newPartPaths, [
    mediaPartAudit.clonePart,
    relationshipPartPath(mediaDuplicateResult.audit.operation.clonePart),
    mediaDuplicateResult.audit.operation.clonePart,
  ].sort());
  assert.equal(mediaDuplicateResult.audit.validation.reimport.sourceAndCloneMediaBindingsIndependent, true);
  assert.deepEqual(await fs.readFile(mediaDuplicateInput), mediaSourceBytes);
  const mediaRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(mediaDuplicateOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "embedded-video-output.pptx",
  }));
  const mediaSourceObject = itemByName(mediaRoundTrip.slides.getItem(0).nativeObjects.items, "Closed embedded video");
  const mediaCloneObject = itemByName(mediaRoundTrip.slides.getItem(1).nativeObjects.items, "Closed embedded video");
  const mediaSourceVideoPart = mediaSourceObject.parts.find((part) => part.contentType === "video/mp4");
  const mediaCloneVideoPart = mediaCloneObject.parts.find((part) => part.contentType === "video/mp4");
  const mediaSourcePosterPart = mediaSourceObject.parts.find((part) => part.contentType.startsWith("image/"));
  const mediaClonePosterPart = mediaCloneObject.parts.find((part) => part.contentType.startsWith("image/"));
  assert.notEqual(mediaSourceVideoPart.path, mediaCloneVideoPart.path);
  assert.equal(mediaSourceVideoPart.sourceSha256, mediaCloneVideoPart.sourceSha256);
  assert.equal(mediaSourcePosterPart.path, mediaClonePosterPart.path);
  const mediaQa = await verifyPresentationFile(mediaDuplicateOutput, {
    outputDir: path.join(duplicateDir, "embedded-video-render-qa"),
    nativeRender,
  });
  assert.equal(mediaQa.verify.ok, true);
  assert.equal(mediaQa.packageInspect.ok, true);
  assert.equal(mediaQa.modelRender.ok, true);
  if (nativeStatus.available) {
    assert.equal(mediaQa.nativeRender.status, "passed");
    assert.deepEqual(
      await fs.readFile(mediaQa.nativeRender.pages[0].path),
      await fs.readFile(mediaQa.nativeRender.pages[1].path),
      "LibreOffice/Poppler must render the shared poster on the source and closed embedded-video clone identically",
    );
  }

  const connectedMediaInput = path.join(duplicateDir, "connected-media-source.pptx");
  const connectedMediaOutput = path.join(duplicateDir, "connected-media-output.pptx");
  const connectedMediaAudit = path.join(duplicateDir, "connected-media-audit.json");
  const connectedMediaZip = await JSZip.loadAsync(mediaSourceBytes);
  connectedMediaZip.file(
    "ppt/media/_rels/skill-video.mp4.rels",
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdUnsafeSkillMedia" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/media" TargetMode="External"/></Relationships>',
  );
  await fs.writeFile(connectedMediaInput, await connectedMediaZip.generateAsync({ type: "nodebuffer" }));
  await assert.rejects(
    () => duplicatePptxSlide({
      inputPath: connectedMediaInput,
      outputPath: connectedMediaOutput,
      auditPath: connectedMediaAudit,
      expectedName: "Embedded video clone source",
    }),
    /Embedded-MP4 payload must not have a child relationship graph/,
  );
  assert.equal(await fs.access(connectedMediaOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(connectedMediaAudit).then(() => true, () => false), false);

  const extensionRichMediaInput = path.join(duplicateDir, "extension-rich-media-source.pptx");
  const extensionRichMediaOutput = path.join(duplicateDir, "extension-rich-media-output.pptx");
  const extensionRichMediaAudit = path.join(duplicateDir, "extension-rich-media-audit.json");
  const extensionRichMediaZip = await JSZip.loadAsync(mediaSourceBytes);
  const extensionRichMediaSlide = await extensionRichMediaZip.file("ppt/slides/slide1.xml").async("text");
  extensionRichMediaZip.file(
    "ppt/slides/slide1.xml",
    extensionRichMediaSlide.replace(
      "</p:extLst>",
      '<p:ext uri="{00000000-0000-0000-0000-000000000000}"><p14:placeholder/></p:ext></p:extLst>',
    ),
  );
  await fs.writeFile(extensionRichMediaInput, await extensionRichMediaZip.generateAsync({ type: "nodebuffer" }));
  await assert.rejects(
    () => duplicatePptxSlide({
      inputPath: extensionRichMediaInput,
      outputPath: extensionRichMediaOutput,
      auditPath: extensionRichMediaAudit,
      expectedName: "Embedded video clone source",
    }),
    /exactly one media action, videoFile, p14:media, poster blip, and canonical media extension/,
  );
  assert.equal(await fs.access(extensionRichMediaOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(extensionRichMediaAudit).then(() => true, () => false), false);

  const connectedInkInput = path.join(duplicateDir, "connected-inkml-source.pptx");
  const connectedInkOutput = path.join(duplicateDir, "connected-inkml-output.pptx");
  const connectedInkAudit = path.join(duplicateDir, "connected-inkml-audit.json");
  const connectedInkZip = await JSZip.loadAsync(inkSourceBytes);
  connectedInkZip.file(
    "ppt/customXml/_rels/skill-ink.xml.rels",
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdUnsafeSkillInk" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/ink" TargetMode="External"/></Relationships>',
  );
  await fs.writeFile(connectedInkInput, await connectedInkZip.generateAsync({ type: "nodebuffer" }));
  await assert.rejects(
    () => duplicatePptxSlide({
      inputPath: connectedInkInput,
      outputPath: connectedInkOutput,
      auditPath: connectedInkAudit,
      expectedName: "InkML clone source",
    }),
    /InkML content part must not have a child relationship graph/,
  );
  assert.equal(await fs.access(connectedInkOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(connectedInkAudit).then(() => true, () => false), false);

  const sharedOleInput = path.join(duplicateDir, "shared-ole-workbook-source.pptx");
  const sharedOleOutput = path.join(duplicateDir, "shared-ole-workbook-output.pptx");
  const sharedOleAudit = path.join(duplicateDir, "shared-ole-workbook-audit.json");
  const sharedOleZip = await JSZip.loadAsync(oleDuplicateSourceBytes);
  const sharedOleRelationships = await sharedOleZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
  sharedOleZip.file("ppt/slides/_rels/slide1.xml.rels", sharedOleRelationships.replace(
    "</Relationships>",
    '<Relationship Id="rIdSharedCloneWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/skill-clone-workbook.xlsx"/></Relationships>',
  ));
  await fs.writeFile(sharedOleInput, await sharedOleZip.generateAsync({ type: "nodebuffer" }));
  await assert.rejects(
    () => duplicatePptxSlide({
      inputPath: sharedOleInput,
      outputPath: sharedOleOutput,
      auditPath: sharedOleAudit,
      expectedName: "OLE clone source",
    }),
    /orphan|exactly one inbound|uniquely bound embedded-XLSX/i,
  );
  assert.equal(await fs.access(sharedOleOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(sharedOleAudit).then(() => true, () => false), false);

  const rootSharedOleInput = path.join(duplicateDir, "root-shared-ole-workbook-source.pptx");
  const rootSharedOleOutput = path.join(duplicateDir, "root-shared-ole-workbook-output.pptx");
  const rootSharedOleAudit = path.join(duplicateDir, "root-shared-ole-workbook-audit.json");
  const rootSharedOleZip = await JSZip.loadAsync(oleDuplicateSourceBytes);
  const rootRelationships = await rootSharedOleZip.file("_rels/.rels").async("text");
  rootSharedOleZip.file("_rels/.rels", rootRelationships.replace(
    "</Relationships>",
    '<Relationship Id="rIdRootSharedCloneWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="ppt/embeddings/skill-clone-workbook.xlsx"/></Relationships>',
  ));
  await fs.writeFile(rootSharedOleInput, await rootSharedOleZip.generateAsync({ type: "nodebuffer" }));
  await assert.rejects(
    () => duplicatePptxSlide({
      inputPath: rootSharedOleInput,
      outputPath: rootSharedOleOutput,
      auditPath: rootSharedOleAudit,
      expectedName: "OLE clone source",
    }),
    /exactly one inbound|uniquely bound embedded-XLSX|unsupported_presentation_slide_clone/i,
  );
  assert.equal(await fs.access(rootSharedOleOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(rootSharedOleAudit).then(() => true, () => false), false);

  const duplicateMissingOutput = path.join(duplicateDir, "missing-target.pptx");
  const duplicateMissingAudit = path.join(duplicateDir, "missing-target.json");
  await assert.rejects(
    () => duplicatePptxSlide({
      inputPath: duplicateInput,
      outputPath: duplicateMissingOutput,
      auditPath: duplicateMissingAudit,
      expectedName: "Missing source slide",
    }),
    /Expected exactly one imported source slide named/,
  );
  assert.equal(await fs.access(duplicateMissingOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(duplicateMissingAudit).then(() => true, () => false), false);

  const orphanLinkInput = path.join(duplicateDir, "orphan-link-source.pptx");
  const orphanLinkOutput = path.join(duplicateDir, "orphan-link-output.pptx");
  const orphanLinkAudit = path.join(duplicateDir, "orphan-link-audit.json");
  const orphanLinkZip = await JSZip.loadAsync(duplicateSourceBytes);
  const orphanRelationships = await orphanLinkZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
  orphanLinkZip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    orphanRelationships.replace(
      "</Relationships>",
      '<Relationship Id="rIdOrphanLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/orphan" TargetMode="External"/></Relationships>',
    ),
  );
  await fs.writeFile(orphanLinkInput, await orphanLinkZip.generateAsync({ type: "nodebuffer" }));
  await assert.rejects(
    () => duplicatePptxSlide({
      inputPath: orphanLinkInput,
      outputPath: orphanLinkOutput,
      auditPath: orphanLinkAudit,
      expectedName: "Clone connector source",
    }),
    /orphan or unmodeled hyperlink relationship/,
  );
  assert.equal(await fs.access(orphanLinkOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(orphanLinkAudit).then(() => true, () => false), false);

  const danglingCustomShowInput = path.join(duplicateDir, "dangling-custom-show-source.pptx");
  const danglingCustomShowOutput = path.join(duplicateDir, "dangling-custom-show-output.pptx");
  const danglingCustomShowAudit = path.join(duplicateDir, "dangling-custom-show-audit.json");
  const danglingCustomShowZip = await JSZip.loadAsync(duplicateSourceBytes);
  const danglingCustomShowSlide = await danglingCustomShowZip.file("ppt/slides/slide1.xml").async("text");
  assert.match(danglingCustomShowSlide, /customshow\?id=17&amp;return=true/);
  danglingCustomShowZip.file(
    "ppt/slides/slide1.xml",
    danglingCustomShowSlide.replace("customshow?id=17&amp;return=true", "customshow?id=999&amp;return=true"),
  );
  await fs.writeFile(danglingCustomShowInput, await danglingCustomShowZip.generateAsync({ type: "nodebuffer" }));
  await assert.rejects(
    () => duplicatePptxSlide({
      inputPath: danglingCustomShowInput,
      outputPath: danglingCustomShowOutput,
      auditPath: danglingCustomShowAudit,
      expectedName: "Clone connector source",
    }),
    /custom-show run action with an unresolved native ID/,
  );
  assert.equal(await fs.access(danglingCustomShowOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(danglingCustomShowAudit).then(() => true, () => false), false);

  const notesInput = path.join(duplicateDir, "notes-source.pptx");
  const notesOutput = path.join(duplicateDir, "notes-output.pptx");
  const notesAudit = path.join(duplicateDir, "notes-audit.json");
  const notesFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
  notesFixture.slides.add({ name: "Source with notes", notes: "This workflow must refuse notes." }).shapes.add({
    name: "notes-copy",
    position: { left: 48, top: 40, width: 260, height: 72 },
    text: "Source with notes",
  });
  await (await PresentationFile.exportPptx(notesFixture)).save(notesInput);
  await assert.rejects(
    () => duplicatePptxSlide({
      inputPath: notesInput,
      outputPath: notesOutput,
      auditPath: notesAudit,
      expectedName: "Source with notes",
    }),
    /accepts no speaker-notes leaf/,
  );
  assert.equal(await fs.access(notesOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(notesAudit).then(() => true, () => false), false);

  const closedLeavesInput = path.join(duplicateDir, "closed-leaves-source.pptx");
  const closedLeavesOutput = path.join(duplicateDir, "closed-leaves-output.pptx");
  const closedLeavesAudit = path.join(duplicateDir, "closed-leaves-audit.json");
  const closedLeavesFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
  const closedLeavesSource = closedLeavesFixture.slides.add({
    name: "Closed leaves source",
    notes: "Open with the decision.\nClose with the owner.",
  });
  closedLeavesSource.shapes.add({
    name: "closed-leaves-title",
    position: { left: 48, top: 40, width: 300, height: 72 },
    text: "Closed leaves source",
  });
  closedLeavesSource.comments.addThread(undefined, "Confirm the original evidence before delivery.", {
    author: "Review Owner",
    created: "2026-07-18T03:05:00Z",
    position: { x: 360, y: 240 },
  });
  await (await PresentationFile.exportPptx(closedLeavesFixture)).save(closedLeavesInput);
  const closedLeavesSourceBytes = await fs.readFile(closedLeavesInput);
  const closedLeavesResult = await duplicatePptxSlide({
    inputPath: closedLeavesInput,
    outputPath: closedLeavesOutput,
    auditPath: closedLeavesAudit,
    expectedName: "Closed leaves source",
    allowClosedLeaves: true,
  });
  assert.equal(closedLeavesResult.audit.operation.scope, "canonical-inline-leaves-with-closed-relationship-leaves");
  assert.deepEqual(closedLeavesResult.audit.operation.closedLeaves, { speakerNotes: true, legacyComments: true });
  const closedLeavesPackage = closedLeavesResult.audit.validation.package;
  const closedNotesAudit = closedLeavesPackage.closedLeaves.speakerNotes;
  const closedCommentsAudit = closedLeavesPackage.closedLeaves.legacyComments;
  assert.deepEqual(closedLeavesResult.audit.validation.package.newPartPaths, [
    closedCommentsAudit.clonePart,
    closedNotesAudit.cloneRelationshipPart,
    closedNotesAudit.clonePart,
    relationshipPartPath(closedLeavesResult.audit.operation.clonePart),
    closedLeavesResult.audit.operation.clonePart,
  ].sort());
  assert.equal(closedLeavesPackage.retainedSourcePartsByteIdentical, true);
  assert.equal(closedNotesAudit.sourcePart, "ppt/notesSlides/notesSlide1.xml");
  assert.match(closedNotesAudit.clonePart, /^ppt\/notesSlides\/[^/]+\.xml$/i);
  assert.notEqual(closedNotesAudit.clonePart, closedNotesAudit.sourcePart);
  assert.equal(closedNotesAudit.sourceRelationshipPart, relationshipPartPath(closedNotesAudit.sourcePart));
  assert.equal(closedNotesAudit.cloneRelationshipPart, relationshipPartPath(closedNotesAudit.clonePart));
  assert.equal(closedNotesAudit.notesMasterPart, "ppt/notesMasters/notesMaster1.xml");
  assert.equal(closedNotesAudit.notesXmlByteIdentical, true);
  assert.equal(closedNotesAudit.notesMasterShared, true);
  assert.equal(closedNotesAudit.cloneBackReferencePointsAtClone, true);
  assert.equal(closedCommentsAudit.sourcePart, "ppt/comments/comment1.xml");
  assert.match(closedCommentsAudit.clonePart, /^ppt\/comments\/[^/]+\.xml$/i);
  assert.notEqual(closedCommentsAudit.clonePart, closedCommentsAudit.sourcePart);
  assert.equal(closedCommentsAudit.commentAuthorsPart, "ppt/commentAuthors.xml");
  assert.equal(closedCommentsAudit.commentsXmlByteIdentical, true);
  assert.equal(closedCommentsAudit.commentAuthorsShared, true);
  assert.equal(closedLeavesResult.audit.validation.reimport.sourceAndCloneClosedLeavesEqual, true);
  assert.deepEqual(await fs.readFile(closedLeavesInput), closedLeavesSourceBytes);
  const closedLeavesRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(closedLeavesOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "closed-leaves-output.pptx",
  }));
  assert.deepEqual(closedLeavesRoundTrip.slides.items.map((slide) => slide.name), ["Closed leaves source", "Closed leaves source"]);
  assert.deepEqual(closedLeavesRoundTrip.slides.items.map((slide) => slide.speakerNotes.text), [
    "Open with the decision.\nClose with the owner.",
    "Open with the decision.\nClose with the owner.",
  ]);
  assert.deepEqual(closedLeavesRoundTrip.slides.items.map((slide) => slide.comments.items[0].comments[0].text), [
    "Confirm the original evidence before delivery.",
    "Confirm the original evidence before delivery.",
  ]);
  await assert.rejects(
    () => duplicatePptxSlide({
      inputPath: closedLeavesInput,
      outputPath: path.join(duplicateDir, "invalid-option.pptx"),
      auditPath: path.join(duplicateDir, "invalid-option.json"),
      expectedName: "Closed leaves source",
      allowClosedLeaves: "yes",
    }),
    /allowClosedLeaves must be a boolean/,
  );

  const irregularLeavesInput = path.join(duplicateDir, "irregular-closed-leaves-source.pptx");
  const irregularLeavesOutput = path.join(duplicateDir, "irregular-closed-leaves-output.pptx");
  const irregularLeavesAudit = path.join(duplicateDir, "irregular-closed-leaves-audit.json");
  const irregularLeavesZip = await JSZip.loadAsync(closedLeavesSourceBytes);
  const sourceNotesRelationships = await irregularLeavesZip.file("ppt/notesSlides/_rels/notesSlide1.xml.rels").async("text");
  irregularLeavesZip.file(
    "ppt/notesSlides/_rels/notesSlide1.xml.rels",
    sourceNotesRelationships.replace(
      "</Relationships>",
      '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="/ppt/customXml/item1.xml" Id="rIdUnexpected" /></Relationships>',
    ),
  );
  irregularLeavesZip.file("ppt/customXml/item1.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><officekit:unexpected xmlns:officekit=\"https://officekit.invalid\" />");
  await fs.writeFile(irregularLeavesInput, await irregularLeavesZip.generateAsync({ type: "nodebuffer" }));
  await assert.rejects(
    () => duplicatePptxSlide({
      inputPath: irregularLeavesInput,
      outputPath: irregularLeavesOutput,
      auditPath: irregularLeavesAudit,
      expectedName: "Closed leaves source",
      allowClosedLeaves: true,
    }),
    /exactly notesMaster and slide relationships/,
  );
  assert.equal(await fs.access(irregularLeavesOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(irregularLeavesAudit).then(() => true, () => false), false);

  const duplicateCliOutput = path.join(duplicateDir, "connector-duplicate-cli.pptx");
  const duplicateCliAudit = path.join(duplicateDir, "cli-audit.json");
  const duplicateCli = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/examples/officekit-slide-duplicate-workflow.mjs",
    duplicateInput,
    duplicateCliOutput,
    duplicateCliAudit,
    "Clone connector source",
  ], { encoding: "utf8" });
  assert.equal(duplicateCli.status, 0, `slide-duplicate CLI failed\n${duplicateCli.stdout}\n${duplicateCli.stderr}`);
  const duplicateCliSummary = JSON.parse(duplicateCli.stdout);
  assert.match(duplicateCliSummary.clonePart, /^ppt\/slides\/slide\d+\.xml$/);
  assert.equal(duplicateCliSummary.customShowActionCount, 1);
  assert.equal(duplicateCliSummary.customShowCount, 1);

  const closedLeavesCliOutput = path.join(duplicateDir, "closed-leaves-cli-output.pptx");
  const closedLeavesCliAudit = path.join(duplicateDir, "closed-leaves-cli-audit.json");
  const closedLeavesCli = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/examples/officekit-slide-duplicate-workflow.mjs",
    closedLeavesInput,
    closedLeavesCliOutput,
    closedLeavesCliAudit,
    "Closed leaves source",
    "--allow-closed-leaves",
  ], { encoding: "utf8" });
  assert.equal(closedLeavesCli.status, 0, `closed-leaves slide-duplicate CLI failed\n${closedLeavesCli.stdout}\n${closedLeavesCli.stderr}`);
  assert.deepEqual(JSON.parse(closedLeavesCli.stdout).closedLeaves, { speakerNotes: true, legacyComments: true });

  const convergenceFiles = [
    "test/skill-harness/presentations/scripts/workflow.mjs",
    "test/skill-harness/presentations/scripts/run-fixture.mjs",
    ...[
      "agent-readiness.json",
      "modern-comments.json",
      "office-kit-legacy-comments.json",
      "office-kit-preservation.json",
      "package-drawing.json",
      "package-notes-comments.json",
    ].map((name) => path.join(fixtureDir, name)),
  ];
  for (const file of convergenceFiles) {
    const source = await fs.readFile(file, "utf8");
    assert.doesNotMatch(source, /\b(?:codec|initialCodec|roundtripCodec)\b/i, file + " must not expose an Office path selector");
  }
  const skillText = `${await fs.readFile("skills/presentations/skills/presentations/SKILL.md", "utf8")}\n${await fs.readFile("skills/presentations/skills/presentations/references/advanced-imported-editing.md", "utf8")}`;
  const conversationWorkflowText = await fs.readFile("skills/presentations/skills/presentations/references/conversation-workflow.md", "utf8");
  const styleGuidelinesText = await fs.readFile("skills/presentations/skills/presentations/style_guidelines.md", "utf8");
  const googleSlidesRoutingText = await fs.readFile("skills/presentations/skills/presentations/routing/google_slides.md", "utf8");
  const quickStartText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/API_QUICK_START.md", "utf8");
  const starterRoot = path.join(root, "template-starter-workflow");
  const starterSource = path.join(starterRoot, "source.pptx");
  const starterMap = path.join(starterRoot, "template-frame-map.json");
  const starterOutput = path.join(starterRoot, "template-starter.pptx");
  const starterManifestPath = path.join(starterRoot, "template-starter.manifest.json");
  const starterPreviewDir = path.join(starterRoot, "preview");
  const starterLayoutDir = path.join(starterRoot, "layout");
  const starterContactSheet = path.join(starterRoot, "contact-sheet.png");
  await fs.mkdir(starterRoot, { recursive: true });
  const starterFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
  for (const [name, text] of [["Alpha", "Alpha body"], ["Beta", "Beta body"], ["Gamma", "Gamma body"]]) {
    const slide = starterFixture.slides.add({ name });
    slide.shapes.add({
      name: `${name}-shape`,
      position: { left: 40, top: 40, width: 300, height: 80 },
      text,
    });
  }
  await (await PresentationFile.exportPptx(starterFixture)).save(starterSource);
  const starterSourceBytes = await fs.readFile(starterSource);
  const inspectStarter = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/template_following_scripts/inspect_template_deck.mjs",
    "--workspace", starterRoot,
    "--pptx", starterSource,
    "--scale", "0.25",
  ], { encoding: "utf8" });
  assert.equal(inspectStarter.status, 0, `template inspection failed\n${inspectStarter.stdout}\n${inspectStarter.stderr}`);
  const starterInspectPath = path.join(starterRoot, "template-inspect", "template-inspect.ndjson");
  const starterInspectRecords = (await fs.readFile(starterInspectPath, "utf8"))
    .split("\n")
    .filter(Boolean)
    .map((line) => JSON.parse(line));
  const betaElementId = starterInspectRecords.find((record) => record.slide === 2 && record.kind !== "slide")?.id;
  assert.equal(typeof betaElementId, "string");
  await fs.writeFile(starterMap, `${JSON.stringify({
    outputSlides: [
      { outputSlide: 1, sourceSlide: 2, narrativeRole: "brand divider", reuseMode: "duplicate-slide", editTargets: [{ action: "keep", sourceElementId: betaElementId }] },
      { outputSlide: 2, sourceSlide: 1, narrativeRole: "brand divider", reuseMode: "duplicate-slide", editTargets: [] },
      { outputSlide: 3, sourceSlide: 2, narrativeRole: "brand divider", reuseMode: "duplicate-slide", editTargets: [{ action: "keep", sourceElementId: betaElementId }] },
    ],
    omittedSourceSlides: [{ sourceSlide: 3, reason: "not needed in the requested narrative" }],
  }, null, 2)}\n`);
  const starterResult = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/template_following_scripts/prepare_template_starter_deck.mjs",
    "--workspace", starterRoot,
    "--pptx", starterSource,
    "--map", starterMap,
    "--out", starterOutput,
    "--preview-dir", starterPreviewDir,
    "--layout-dir", starterLayoutDir,
    "--contact-sheet", starterContactSheet,
    "--scale", "0.25",
  ], { encoding: "utf8" });
  assert.equal(starterResult.status, 0, `template starter failed\n${starterResult.stdout}\n${starterResult.stderr}`);
  assert.ok(JSON.parse(starterResult.stdout).outputSha256);
  assert.deepEqual(await fs.readFile(starterSource), starterSourceBytes, "starter generation must not mutate the source PPTX");
  const starterPresentation = await PresentationFile.importPptx(await FileBlob.load(starterOutput));
  assert.deepEqual(starterPresentation.slides.items.map((slide) => slide.name), ["Beta", "Alpha", "Beta"]);
  assert.equal(starterPresentation.verify({ visualQa: true }).ok, true);
  const starterManifest = JSON.parse(await fs.readFile(starterManifestPath, "utf8"));
  assert.equal(starterManifest.schema, "office-kit.template-starter.v1");
  assert.equal(starterManifest.status, "succeeded");
  assert.deepEqual(starterManifest.operation.cloneBoundaries.map((entry) => entry.sourceSlide), [2, 1, 2]);
  assert.equal(starterManifest.operation.cloneBoundaries.every((entry) => entry.exportReimported === true), true);
  assert.equal(starterManifest.validation.oneClonePerExportBoundary, true);
  assert.equal(starterManifest.validation.finalExportReimported, true);
  assert.equal(starterManifest.validation.locatorTranslationComplete, true);
  assert.deepEqual(starterManifest.slides.map((entry) => entry.sourceSlide), [2, 1, 2]);
  const firstBetaTarget = starterManifest.slides[0].editTargets[0];
  const secondBetaTarget = starterManifest.slides[2].editTargets[0];
  assert.deepEqual(firstBetaTarget.sourceElementIds, [betaElementId]);
  assert.deepEqual(secondBetaTarget.sourceElementIds, [betaElementId]);
  assert.equal(firstBetaTarget.starterElementIds.length, 1);
  assert.equal(secondBetaTarget.starterElementIds.length, 1);
  assert.notEqual(firstBetaTarget.starterElementIds[0], secondBetaTarget.starterElementIds[0]);
  assert.deepEqual((await fs.readdir(starterPreviewDir)).sort(), [
    "starter-slide-01.png",
    "starter-slide-02.png",
    "starter-slide-03.png",
  ]);
  assert.deepEqual((await fs.readdir(starterLayoutDir)).sort(), [
    "starter-slide-01.layout.json",
    "starter-slide-02.layout.json",
    "starter-slide-03.layout.json",
  ]);
  assert.deepEqual([...await fs.readFile(starterContactSheet)].slice(0, 8), [137, 80, 78, 71, 13, 10, 26, 10]);

  const starterOutputBytes = await fs.readFile(starterOutput);
  const starterOverwrite = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/template_following_scripts/prepare_template_starter_deck.mjs",
    "--workspace", starterRoot,
    "--pptx", starterSource,
    "--map", starterMap,
    "--out", starterOutput,
    "--preview-dir", starterPreviewDir,
    "--layout-dir", starterLayoutDir,
    "--contact-sheet", starterContactSheet,
  ], { encoding: "utf8" });
  assert.notEqual(starterOverwrite.status, 0);
  assert.match(starterOverwrite.stderr, /already exists; refusing to overwrite/i);
  assert.deepEqual(await fs.readFile(starterOutput), starterOutputBytes);
  assert.deepEqual(await fs.readFile(starterSource), starterSourceBytes);

  const invalidStarterRoot = path.join(root, "starter-invalid-plan");
  const invalidStarterMap = path.join(invalidStarterRoot, "template-frame-map.json");
  const invalidStarterOutput = path.join(invalidStarterRoot, "template-starter.pptx");
  await fs.mkdir(invalidStarterRoot, { recursive: true });
  await fs.writeFile(invalidStarterMap, JSON.stringify({ outputSlides: [] }));
  const invalidStarter = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/template_following_scripts/prepare_template_starter_deck.mjs",
    "--workspace", invalidStarterRoot,
    "--pptx", starterSource,
    "--map", invalidStarterMap,
    "--out", invalidStarterOutput,
  ], { encoding: "utf8" });
  assert.notEqual(invalidStarter.status, 0);
  assert.match(invalidStarter.stderr, /Missing template inspection/i);
  assert.deepEqual((await fs.readdir(invalidStarterRoot)).sort(), ["template-frame-map.json"]);

  const blockedStarterRoot = path.join(root, "starter-blocked-graph");
  const blockedStarterSource = path.join(blockedStarterRoot, "source.pptx");
  const blockedStarterMap = path.join(blockedStarterRoot, "template-frame-map.json");
  const blockedStarterOutput = path.join(blockedStarterRoot, "template-starter.pptx");
  const blockedStarterPreview = path.join(blockedStarterRoot, "preview");
  const blockedStarterLayout = path.join(blockedStarterRoot, "layout");
  await fs.mkdir(blockedStarterRoot, { recursive: true });
  const linkedStarterFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
  const linkedStarterOrigin = linkedStarterFixture.slides.add({ name: "Linked origin" });
  const linkedStarterTarget = linkedStarterFixture.slides.add({ name: "Linked target" });
  linkedStarterTarget.shapes.add({ name: "Target", position: { left: 40, top: 40, width: 300, height: 80 }, text: "Target" });
  linkedStarterOrigin.shapes.add({
    name: "Jump",
    position: { left: 40, top: 40, width: 300, height: 80 },
    text: [{ runs: [{ text: "Jump to target", link: { slideId: linkedStarterTarget.id } }] }],
  });
  await (await PresentationFile.exportPptx(linkedStarterFixture)).save(blockedStarterSource);
  const blockedStarterSourceBytes = await fs.readFile(blockedStarterSource);
  const blockedInspect = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/template_following_scripts/inspect_template_deck.mjs",
    "--workspace", blockedStarterRoot,
    "--pptx", blockedStarterSource,
    "--scale", "0.25",
  ], { encoding: "utf8" });
  assert.equal(blockedInspect.status, 0, `blocked template inspection failed\n${blockedInspect.stdout}\n${blockedInspect.stderr}`);
  await fs.writeFile(blockedStarterMap, `${JSON.stringify({
    outputSlides: [{ outputSlide: 1, sourceSlide: 1, narrativeRole: "brand divider", reuseMode: "duplicate-slide", editTargets: [] }],
    omittedSourceSlides: [{ sourceSlide: 2, reason: "target slide omitted" }],
  }, null, 2)}\n`);
  const blockedStarter = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/template_following_scripts/prepare_template_starter_deck.mjs",
    "--workspace", blockedStarterRoot,
    "--pptx", blockedStarterSource,
    "--map", blockedStarterMap,
    "--out", blockedStarterOutput,
    "--preview-dir", blockedStarterPreview,
    "--layout-dir", blockedStarterLayout,
  ], { encoding: "utf8" });
  assert.notEqual(blockedStarter.status, 0);
  assert.match(blockedStarter.stderr, /Cannot remove sourceSlide 2[\s\S]*referenced/i);
  assert.deepEqual(await fs.readFile(blockedStarterSource), blockedStarterSourceBytes);
  assert.equal(await fs.access(blockedStarterOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(path.join(blockedStarterRoot, "template-starter.manifest.json")).then(() => true, () => false), false);
  assert.equal(await fs.access(blockedStarterPreview).then(() => true, () => false), false);
  assert.equal(await fs.access(blockedStarterLayout).then(() => true, () => false), false);
  assert.equal((await fs.readdir(blockedStarterRoot)).some((name) => name.startsWith(".office-kit-template-starter-")), false);
  assert.match(skillText, /office-kit/);
  assert.match(skillText, /conversational workflow/i);
  assert.match(skillText, /working draft.*ctx\.publish.*after acceptance/is);
  assert.match(skillText, /ask at most three questions/i);
  assert.match(skillText, /one-screen goal\/structure\/confirmation guidance/i);
  assert.match(conversationWorkflowText, /net-new deck or broad redesign/i);
  assert.match(conversationWorkflowText, /one-screen draft guide/i);
  assert.match(conversationWorkflowText, /Never reuse a stale locator/i);
  assert.match(conversationWorkflowText, /Explicit .* permits complete\s+post-edit review and final publication/is);
  assert.doesNotMatch(conversationWorkflowText, /(?:invoke|use)\s+\$presentations/i);
  assert.match(skillText, /officekit-speaker-notes-add-workflow\.mjs/);
  assert.match(skillText, /officekit-rich-speaker-notes-edit-workflow\.mjs/);
  assert.match(skillText, /paragraph `0`, run `1`[\s\S]*not widen the topology/is);
  assert.match(skillText, /### Rich Speaker Notes/);
  assert.match(skillText, /speakerNotes\.capability\.addable.*existing.*NotesMaster.*byte-for-byte.*canonical NotesMaster.*ThemePart.*back-reference/is);
  assert.match(skillText, /officekit-legacy-comment-add-workflow\.mjs/);
  assert.match(skillText, /officekit-legacy-comment-add-workflow\.mjs/);
  assert.match(skillText, /comments\.capability.*format: "legacy".*partPresent: false.*editable: false.*addable: true.*CommentAuthorsPart.*SlideCommentsPart.*collision-free.*pixel-identical/is);
  assert.match(skillText, /officekit-legacy-comment-edit-workflow\.mjs/);
  assert.match(skillText, /comments\.capability.*format: "legacy"?, partPresent: true, editable: true/is);
  assert.match(skillText, /stable comment ID.*expected\s+old text/is);
  assert.match(skillText, /only the selected root text.*author catalog.*byte-identical.*ppt\/comments\/commentN\.xml.*fail\s+closed/is);
  assert.match(skillText, /officekit-title-notes-edit-workflow\.mjs/);
  assert.match(skillText, /officekit-modern-comment-workflow\.mjs/);
  assert.match(skillText, /officekit-slide-name-edit-workflow\.mjs/);
  assert.match(skillText, /Slide Show Visibility/);
  assert.match(skillText, /visibilityCapability\.known.*visibilityCapability\.editable.*slide\.hide\(\).*slide\.show\(\).*p:sld\/@show="0".*fail closed/is);
  assert.match(skillText, /officekit-transition-edit-workflow\.mjs/);
  assert.match(skillText, /officekit-slide-duplicate-workflow\.mjs/);
  assert.match(skillText, /table\.getCell\(row, column\)\.value.*ImageElement\.dataUrl.*deletionCapability.*tableCellText.*imageAsset.*deleteElement.*do not synthesize native leaf IDs.*package footprint.*unchanged comparison pages/is);
  assert.match(skillText, /presentation\.auditAccessibility\(\)/);
  assert.match(skillText, /officekit-accessibility-audit-workflow\.mjs/);
  assert.match(skillText, /machineCheckPassed: true.*whole-deck conformance/is);
  assert.match(skillText, /--allow-closed-leaves/);
  assert.match(quickStartText, /PresentationFile\.exportPptx/);
  assert.match(quickStartText, /addPptxSpeakerNotes/);
  assert.match(quickStartText, /editPptxRichSpeakerNotes/);
  assert.match(quickStartText, /fixed-topology transaction validates the imported source run text/i);
  assert.match(quickStartText, /notes\.capability\.sourceBound.*notes\.capability\.partPresent.*notes\.capability\.addable/is);
  assert.match(quickStartText, /addPptxLegacyReviewComment/);
  assert.match(quickStartText, /editPptxLegacyReviewComment/);
  assert.match(quickStartText, /comments\.capability.*sourceBound.*format.*partPresent.*editable.*addable.*no legacy or Office 2021 comment graph.*re-proves/is);
  assert.match(quickStartText, /editPptxSlideName/);
  assert.match(quickStartText, /slide\.setHidden\(true\).*slide\.show\(\).*ordinary\s+slide-show playback/is);
  assert.match(quickStartText, /officekit-transition-edit-workflow\.mjs/);
  assert.match(quickStartText, /duplicatePptxSlide/);
  assert.match(quickStartText, /allowClosedLeaves:\s*true/);
  assert.match(quickStartText, /commentFormat:\s*"modern"/);
  assert.match(quickStartText, /office-kit/);
  assert.match(skillText, /slides_test\.py/);
  assert.match(skillText, /slide\.setBackground.*slide\.clearBackground/s);
  assert.match(skillText, /complete ECMA-376 base.*transition vocabulary/is);
  assert.match(skillText, /slide\.setTransition\(\{.*effect: "split".*orientation: "horizontal".*direction: "in".*advanceOnClick.*advanceAfterMs/is);
  assert.match(skillText, /transition\.capability.*canonical direct base-transition profile.*no transition.*addable: true.*p:cSld.*p:clrMapOvr.*no transition, timing, or extension leaf.*timing.*sound.*p14.*extension.*opaque-preserved.*fail closed/is);
  assert.match(skillText, /slide\.moveTo\(existingZeroBasedIndex\).*retained source.*p:sldIdLst.*slide\.deletionCapability.*exclusively owned OPC descendant.*shared layout\/master\/theme\/image\/media.*fail closed/is);
  assert.match(skillText, /Top-level imported ordinary shapes.*embedded pictures.*canonical connectors.*bounded tables.*charts.*deletionCapability.*relationship-free subtree.*ImagePart\/ChartPart descendant closure.*shared media and ChartParts remain.*array splicing/is);
  assert.match(skillText, /starter command executes a validated multi-slide frame map/is);
  assert.match(skillText, /duplicates exactly one supported source slide per export\/reimport boundary.*same source slide may be reused safely.*removes all\s+original slides.*ownership transaction/is);
  assert.match(skillText, /translates every inspected source element ID.*final starter ID.*fails closed/is);
  assert.match(skillText, /slide\.duplicate\(\).*slide\.cloneCapability.*OPC ownership graph.*recursively copies.*uniquely owned OpenXmlPart.*DataPart.*exact part bytes.*external relationships.*rebinds only proven shared layouts.*NotesMaster.*images.*retained slide-jump targets.*Unknown or relationship-bearing descendants.*not rejected/is);
  assert.match(skillText, /Recognized closed ChartParts.*one unique frame relationship.*no ChartPart child graph.*distinct clone-local target.*byte-identical chart payload/is);
  assert.match(skillText, /eligible embedded-XLSX OLE workbook.*one unique inbound package edge.*empty child graph.*same slide-local `r:id`.*distinct clone package.*shared preview\s+ImagePart/is);
  assert.match(skillText, /accepted InkML content part.*one exact `customXml`\s+relationship.*standard content type and root namespace.*empty child\s+graph.*distinct clone-local `CustomXmlPart`.*byte-identical XML/is);
  assert.match(skillText, /video\/media relationship pairing.*unique inbound ownership.*exact\s+`video\/mp4` bytes.*distinct clone-local `MediaDataPart`.*shared immutable\s+poster/is);
  assert.match(skillText, /bounded clone workflow.*canonical relationship-free run action.*show membership\s+did not change/is);
  assert.match(quickStartText, /recognized literal-data charts.*no child\/external\/hyperlink\/data relationship.*distinct byte-copied ChartPart/is);
  assert.match(quickStartText, /eligible top-level OLE frames.*uniquely inbound internal XLSX package.*distinct\s+byte-copied EmbeddedPackagePart.*shares only the immutable preview/is);
  assert.match(quickStartText, /canonical top-level SmartArt frames.*data\/layout\/\s*quick-style\/colors parts.*four\s+distinct typed diagram parts/is);
  assert.match(quickStartText, /canonical top-level `p:contentPart`.*closed standard InkML part.*distinct byte-identical SDK `CustomXmlPart`/is);
  assert.match(quickStartText, /embedded-MP4.*video\/media relationships.*distinct byte-identical SDK `MediaDataPart`.*shares only the immutable preview/is);
  assert.match(quickStartText, /relationship-free custom-show action.*exact native ID\/return policy.*clone.*not silently added to the route/is);
  assert.match(skillText, /NotesSlide.*NotesMaster.*byte-for-byte.*back-reference.*clone/is);
  assert.match(skillText, /SlideCommentsPart.*CommentAuthorsPart.*byte-for-byte/is);
  assert.match(skillText, /artifact_tool\/api\/references\/comments\.md/);
  assert.match(skillText, /Presentation Style Guidelines/);
  assert.match(skillText, /routing\/google_slides\.md/);
  assert.match(styleGuidelinesText, /By the end, \*\*\[audience\]\*\* should \*\*\[outcome\]\*\*/);
  assert.match(styleGuidelinesText, /Never invent people, quotes, facts, data, or outcomes/);
  assert.match(styleGuidelinesText, /50 pt for deck titles/);
  assert.match(googleSlidesRoutingText, /local `\.pptx`/);
  assert.match(googleSlidesRoutingText, /separate host\s+step/);
  assert.doesNotMatch(googleSlidesRoutingText, /plugin|mcp__|google-drive@/i);
  const speakerNotesReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/speaker-notes.spec.md", "utf8");
  assert.match(speakerNotesReferenceText, /sourceBound.*partPresent.*editable.*addable/is);
  assert.match(speakerNotesReferenceText, /relationship-free.*paragraph\/run/is);
  assert.match(speakerNotesReferenceText, /editPptxRichSpeakerNotes/);
  assert.match(speakerNotesReferenceText, /not a general NotesSlide\s+reflow/i);
  assert.match(speakerNotesReferenceText, /fields.*hyperlinks.*picture bullets.*fail(?:s)?\s+closed/is);
  assert.match(speakerNotesReferenceText, /capability is preflight evidence.*not authority.*independently re-proves/is);
  assert.match(speakerNotesReferenceText, /existing.*NotesMaster.*reused byte-for-byte.*canonical NotesMaster.*ThemePart.*back-reference/is);
  const commentsReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/comments.md", "utf8");
  assert.match(commentsReferenceText, /Pass `undefined` as the target/);
  assert.match(commentsReferenceText, /one author, one text item, and one explicit\s+slide coordinate/is);
  assert.match(commentsReferenceText, /Office 2021 modern threads/);
  assert.match(commentsReferenceText, /Only existing comment text and status are mutable/);
  assert.match(commentsReferenceText, /Reactions\/likes, task fields, extensions, rich text, nested replies/);
  assert.match(commentsReferenceText, /officekit-modern-comment-workflow\.mjs/);
  assert.match(commentsReferenceText, /slide\.duplicate\(\).*byte-cop(?:y|ied).*author\s+catalog/is);
  const slideReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/slide.spec.md", "utf8");
  assert.match(slideReferenceText, /slide\.visibilityCapability.*slide\.setHidden\(true\).*slide\.hide\(\).*slide\.show\(\)/s);
  assert.match(slideReferenceText, /p:sld\/@show="0".*invalid.*opaque.*fail/s);
  assert.match(slideReferenceText, /never flattens the\s+inherited color/i);
  assert.match(slideReferenceText, /NotesSlide.*NotesMaster.*exactly those two\s+relationships.*byte-for-byte/is);
  assert.match(slideReferenceText, /canonical inline fixed-grid tables[\s\S]*cannot introduce a fill, link/i);
  assert.match(slideReferenceText, /recognized closed literal-data charts.*numbered\s+`ChartPart`.*no child, external, hyperlink, or data relationship.*distinct clone-local ChartPart/is);
  assert.match(slideReferenceText, /eligible top-level embedded-XLSX OLE frames.*uniquely binds one\s+closed, uniquely inbound internal XLSX.*distinct clone-local package.*replaceEmbeddedWorkbook/is);
  assert.match(slideReferenceText, /SmartArt frame.*dgm:relIds.*dm\/lo\/qs\/cs.*distinct clone-local typed parts.*source-bound\/read-only/is);
  assert.match(slideReferenceText, /InkML object.*top-level `p:contentPart`.*`application\/inkml\+xml`.*distinct SDK `CustomXmlPart`.*source-bound\/read-only/is);
  assert.match(slideReferenceText, /accepted embedded video.*top-level `p:pic`.*`p14:media`.*distinct SDK `MediaDataPart`.*shares the immutable poster.*playback/is);
  assert.match(slideReferenceText, /Gradient,\s+pattern, image.*opaque-preserved/is);
  assert.match(slideReferenceText, /p:cSld\/@name.*export\/reimport/is);
  assert.match(slideReferenceText, /direct transition profile.*complete ECMA-376 base.*absent transition may be added only when.*transition\.capability\.addable.*`p:cSld`.*`p:clrMapOvr`.*opaque-preserved/is);
  const transitionReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/transitions.spec.md", "utf8");
  assert.match(transitionReferenceText, /`p:transition` contract.*not a PowerPoint timing or\s+animation engine/is);
  assert.match(transitionReferenceText, /21 base effects.*`blinds`.*`wheel`.*`zoom`/is);
  assert.match(transitionReferenceText, /`orientation`.*horizontal.*vertical.*`throughBlack`.*`spokes`.*1\.\.8/is);
  assert.match(transitionReferenceText, /durationMs.*advanceOnClick.*advanceAfterMs.*0\.\.86400000/is);
  assert.match(transitionReferenceText, /with no transition is addable only when.*capability\.addable.*`p:cSld`.*`p:clrMapOvr`.*`p:transition`, `p:timing`, or extension leaf/is);
  assert.match(transitionReferenceText, /non-integer.*`p14:dur`.*`p:timing`.*sound.*extension.*opaque.*byte-for-byte/is);
  assert.match(transitionReferenceText, /static PNG\/PDF\s+render cannot prove slideshow playback/is);
  assert.match(transitionReferenceText, /officekit-transition-edit-workflow\.mjs.*unique imported slide name.*expected source state.*replacement state/is);
  assert.match(transitionReferenceText, /non-target parts.*Exactly the selected SlidePart must differ/is);
  const customShowReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/custom-shows.spec.md", "utf8");
  assert.match(customShowReferenceText, /bounded slide clone profile.*relationship-free action.*creates no hyperlink\/slide relationship.*clone is not implicitly added/is);
  const imageReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/images.spec.md", "utf8");
  assert.match(imageReferenceText, /signed normalized source edges in `-1\.\.1`/i);
  assert.match(imageReferenceText, /DrawingML `a:srcRect`/);
  assert.match(imageReferenceText, /PPTX has no native fit keyword/i);
  assert.match(imageReferenceText, /unsafe edits fail\s+closed/i);
  const accessibilityReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/accessibility.spec.md", "utf8");
  assert.match(accessibilityReferenceText, /meaningful ordinary shapes, connectors, groups, images, tables, and charts/i);
  assert.match(accessibilityReferenceText, /conformanceClaimed: false/);
  assert.match(accessibilityReferenceText, /shape tree would also change\s+visual z-order/i);
  assert.match(accessibilityReferenceText, /officekit-accessibility-audit-workflow\.mjs/);
  assert.match(accessibilityReferenceText, /source mutation.*fail(?:s)? closed/is);
  assert.match(skillText, /complete audit locator.*officekit-object-accessibility-edit-workflow\.mjs.*exactly the selected SlidePart/is);
  assert.match(accessibilityReferenceText, /officekit-object-accessibility-edit-workflow\.mjs/);
  assert.match(accessibilityReferenceText, /shape, connector,\s+group, image, table, or chart/is);
  assert.match(accessibilityReferenceText, /exactly the\s+selected SlidePart to differ.*complete non-target presentation projection/is);
  assert.match(accessibilityReferenceText, /not whole-deck\s+accessibility conformance/is);
  const groupingReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/grouping.spec.md", "utf8");
  assert.match(skillText, /artifact_tool\/api\/references\/grouping\.spec\.md/);
  assert.match(groupingReferenceText, /real `p:grpSp`/);
  assert.match(groupingReferenceText, /a:chOff\/a:chExt/);
  assert.match(groupingReferenceText, /presentation_group_topology_changed/);
  assert.match(groupingReferenceText, /one opaque, read-only native object/i);
  const oleWorkbookReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/ole-workbooks.spec.md", "utf8");
  assert.match(skillText, /artifact_tool\/api\/references\/ole-workbooks\.spec\.md/);
  assert.match(oleWorkbookReferenceText, /getEmbeddedWorkbook\(\).*replaceEmbeddedWorkbook/s);
  assert.match(oleWorkbookReferenceText, /getEmbeddedOfficePackage\(\).*replaceEmbeddedOfficePackage/s);
  assert.match(oleWorkbookReferenceText, /only newly supported kind is a DOCX package/i);
  assert.match(oleWorkbookReferenceText, /not a generic OLE\/container API/i);
  assert.match(oleWorkbookReferenceText, /officekit-ole-office-package-workflow\.mjs/);
  assert.match(skillText, /getEmbeddedOfficePackage\(\).*replaceEmbeddedOfficePackage\(\.\.\.\).*DOCX/is);
  assert.match(oleWorkbookReferenceText, /preserving the OLE\s+shell, relationship topology, preview image/is);
  assert.match(oleWorkbookReferenceText, /Microsoft Open XML\s+SDK/);
  assert.match(oleWorkbookReferenceText, /shared, external, ambiguous, or unsupported package/);
  assert.match(oleWorkbookReferenceText, /slide\.duplicate\(\).*distinct clone-local XLSX `EmbeddedPackagePart`.*shares\s+the immutable preview ImagePart.*export has been imported again/is);
  assert.match(oleWorkbookReferenceText, /no lossy reconstruction or silent\s+fallback/i);
  const smartArtReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/smartart-clone.spec.md", "utf8");
  assert.match(skillText, /artifact_tool\/api\/references\/smartart-clone\.spec\.md/);
  assert.match(skillText, /officekit-smartart-text-edit-workflow\.mjs/);
  assert.match(smartArtReferenceText, /top-level `p:graphicFrame`.*exactly one `dgm:relIds`/is);
  assert.match(smartArtReferenceText, /four distinct typed diagram parts.*disjoint part paths.*per-role hashes/is);
  assert.match(smartArtReferenceText, /Neither contract is SmartArt\s+authoring.*fail closed/is);
  assert.match(smartArtReferenceText, /dgm:dataModel.*dgm:ptLst.*32,767/is);
  assert.match(smartArtReferenceText, /one\s+through 256 direct `dgm:t > a:p` paragraphs.*between one and 256\s+direct `a:r > a:t` runs plus fixed line breaks in total.*Empty direct paragraphs.*wholly empty node/is);
  assert.match(smartArtReferenceText, /setDiagramNodeRunText.*runs.*source order.*a:pPr.*a:rPr.*a:br.*a:endParaRPr.*--run-index=1/is);
  assert.match(smartArtReferenceText, /only the DiagramDataPart changed.*reimports the graph.*LibreOffice\/Poppler/is);
  const inkMlReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/inkml-content-part-clone.spec.md", "utf8");
  assert.match(skillText, /artifact_tool\/api\/references\/inkml-content-part-clone\.spec\.md/);
  assert.match(inkMlReferenceText, /top-level child.*`p:spTree`.*exactly one relationship attribute/is);
  assert.match(inkMlReferenceText, /`application\/inkml\+xml`.*`http:\/\/www\.w3\.org\/2003\/InkML`/is);
  assert.match(inkMlReferenceText, /distinct Open\s+XML SDK `CustomXmlPart`.*disjoint part paths.*equal payload hashes/is);
  assert.match(inkMlReferenceText, /opaque and read-only.*fail closed/is);
  const embeddedVideoReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/embedded-video-clone.spec.md", "utf8");
  assert.match(skillText, /artifact_tool\/api\/references\/embedded-video-clone\.spec\.md/);
  assert.match(embeddedVideoReferenceText, /top-level `p:pic`.*empty `r:id` media-action sentinel.*`a:videoFile\/@r:link`.*`p14:media\/@r:embed`/is);
  assert.match(embeddedVideoReferenceText, /distinct `MediaDataPart`.*exact MP4 bytes.*same poster path/is);
  assert.match(embeddedVideoReferenceText, /poster remains equal.*do not claim media playback equivalence/is);
  assert.match(embeddedVideoReferenceText, /audio.*linked or\s+external media.*timing.*fail closed/is);
  const templateFollowingText = await fs.readFile("skills/presentations/skills/presentations/references/template-following.md", "utf8");
  assert.match(templateFollowingText, /slide\.cloneCapability.*duplicates one.*imported slide's closed, uniquely owned OPC descendant graph.*unknown parts and external relationships.*multi-slide frame map.*one clone per export\/reimport boundary.*repeated.*source slide.*deletion-capability preflight.*source-to-starter locators/is);
  assert.match(templateFollowingText, /Do not substitute a reconstructed or shared-part copy.*Unsupported topology.*fails closed.*before any starter artifact is published/is);
  assert.match(templateFollowingText, /independently revalidates the map.*Every clone is exported and imported.*before the next clone.*removed in reverse order.*one-pending-\s*clone contract.*no PPTX, manifest, preview,\s+layout, or contact-sheet publication/is);
  assert.match(templateFollowingText, /template-starter\.manifest\.json.*sourceElementIds.*starterElementIds.*do not claim persistence.*source\/map\/inspection\/output hashes.*no-overwrite policy/is);

  console.log("presentation skill smoke ok");
} finally {
  await fs.rm(root, { recursive: true, force: true });
}
