import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { FileBlob, Presentation, PresentationFile } from "../src/index.mjs";
import { applyTemplateEditPlan } from "../skills/presentations/skills/presentations/template_following_scripts/apply_template_edit_plan.mjs";
import { prepareTemplateStarterDeck } from "../skills/presentations/skills/presentations/template_following_scripts/prepare_template_starter_deck.mjs";

const ORIGINAL_PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
const REPLACEMENT_PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function dataUrlBytes(value) {
  return Buffer.from(String(value).split(",", 2)[1], "base64");
}

async function writeJson(filePath, value) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

function targetByIndex(manifest, outputSlide, targetIndex) {
  return manifest.slides[outputSlide - 1].editTargets.find((target) => target.targetIndex === targetIndex);
}

const [skillText, referenceText] = await Promise.all([
  fs.readFile("skills/presentations/skills/presentations/SKILL.md", "utf8"),
  fs.readFile("skills/presentations/skills/presentations/references/template-following.md", "utf8"),
]);
assert.match(skillText, /apply_template_edit_plan\.mjs[\s\S]*template-edit-plan\.json[\s\S]*publishes nothing/i);
assert.match(referenceText, /office-kit\.template-edit-plan\.v1[\s\S]*set-text[\s\S]*replace-text[\s\S]*set-table-cell[\s\S]*set-chart-series-values[\s\S]*replace-image/i);
assert.match(referenceText, /does not currently execute `delete`[\s\S]*Do not claim a\s+placeholder was removed/i);

const root = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-template-edit-"));
try {
  const sourcePath = path.join(root, "source.pptx");
  const inspectPath = path.join(root, "template-inspect", "template-inspect.ndjson");
  const mapPath = path.join(root, "template-frame-map.json");
  const starterPath = path.join(root, "template-starter.pptx");
  const starterManifestPath = path.join(root, "template-starter.manifest.json");
  const starterPreviewDir = path.join(root, "starter-preview");
  const starterLayoutDir = path.join(root, "starter-layout");
  const assetPath = path.join(root, "replacement.png");
  const planPath = path.join(root, "template-edit-plan.json");
  const outputPath = path.join(root, "deliverables", "final.pptx");
  const auditPath = path.join(root, "deliverables", "final.audit.json");
  const previewDir = path.join(root, "final-preview");
  const layoutDir = path.join(root, "final-layout");
  const contactSheetPath = path.join(root, "final-contact-sheet.png");

  const source = Presentation.create({ slideSize: { width: 640, height: 360 } });
  const editSlide = source.slides.add({ name: "Editable frame" });
  editSlide.shapes.add({
    name: "title-target",
    position: { left: 40, top: 20, width: 260, height: 90 },
    text: [{ runs: [{ text: "Revenue outlook pending", style: { bold: true, fontSize: 28 } }] }],
  });
  editSlide.shapes.add({
    name: "owner-target",
    position: { left: 40, top: 240, width: 220, height: 70 },
    text: [{ runs: [{ text: "Owner pending", style: { fontSize: 18 } }] }],
  });
  editSlide.tables.add({
    name: "table-target",
    rows: 2,
    columns: 2,
    position: { left: 40, top: 125, width: 220, height: 100 },
    values: [["Quarter", "Revenue"], ["Q1", "120"]],
  });
  editSlide.charts.add("bar", {
    name: "chart-target",
    title: "Revenue",
    position: { left: 290, top: 110, width: 300, height: 180 },
    categories: ["Q1", "Q2"],
    series: [{ name: "Plan", values: [120, 130] }],
  });
  editSlide.images.add({
    name: "image-target",
    alt: "Product mark",
    position: { left: 500, top: 24, width: 72, height: 72 },
    dataUrl: ORIGINAL_PNG,
    fit: "contain",
  });
  const preservedSlide = source.slides.add({ name: "Preserved frame" });
  preservedSlide.shapes.add({
    name: "preserve-target",
    position: { left: 60, top: 80, width: 500, height: 120 },
    text: "This slide must remain unchanged.",
  });
  await (await PresentationFile.exportPptx(source)).save(sourcePath);
  const sourceBytes = await fs.readFile(sourcePath);
  const importedSource = await PresentationFile.importPptx(await FileBlob.load(sourcePath));
  const inspection = importedSource.inspect({ kind: "slide,textbox,shape,image,table,chart,connector,groupShape,nativeObject", maxChars: 2_000_000 });
  assert.equal(inspection.truncated, false);
  await fs.mkdir(path.dirname(inspectPath), { recursive: true });
  await fs.writeFile(inspectPath, inspection.ndjson, "utf8");
  const byName = new Map(inspection.ndjson.split("\n").filter(Boolean).map((line) => JSON.parse(line)).filter((record) => record.name).map((record) => [record.name, record.id]));
  await writeJson(mapPath, {
    outputSlides: [
      {
        outputSlide: 1,
        sourceSlide: 1,
        narrativeRole: "revenue analysis",
        reuseMode: "duplicate-slide",
        editTargets: [
          { action: "rewrite-and-reposition", sourceElementId: byName.get("title-target") },
          { action: "rewrite", sourceElementId: byName.get("owner-target") },
          { action: "rewrite", sourceElementId: byName.get("table-target") },
          { action: "rewrite", sourceElementId: byName.get("chart-target") },
          { action: "replace", sourceElementId: byName.get("image-target") },
        ],
      },
      {
        outputSlide: 2,
        sourceSlide: 2,
        narrativeRole: "brand divider",
        reuseMode: "duplicate-slide",
        editTargets: [{ action: "keep", sourceElementId: byName.get("preserve-target") }],
      },
    ],
  });
  await prepareTemplateStarterDeck({
    workspace: root,
    pptxPath: sourcePath,
    mapPath,
    inspectPath,
    out: starterPath,
    previewDir: starterPreviewDir,
    layoutDir: starterLayoutDir,
    scale: 0.25,
  });
  const starterBytes = await fs.readFile(starterPath);
  const starterManifestBytes = await fs.readFile(starterManifestPath);
  const starterManifest = JSON.parse(starterManifestBytes);
  const replacementBytes = dataUrlBytes(REPLACEMENT_PNG);
  await fs.writeFile(assetPath, replacementBytes);
  const plan = {
    schema: "office-kit.template-edit-plan.v1",
    starterSha256: sha256(starterBytes),
    manifestSha256: sha256(starterManifestBytes),
    targets: [
      {
        outputSlide: 1,
        targetIndex: 0,
        operations: [
          { type: "replace-text", expectedText: "Revenue outlook pending", text: "Revenue outlook approved" },
          {
            type: "set-position",
            expectedPosition: { left: 40, top: 20, width: 260, height: 90 },
            position: { left: 48, top: 18, width: 270, height: 88 },
          },
        ],
      },
      {
        outputSlide: 1,
        targetIndex: 1,
        operations: [{ type: "set-text", expectedText: "Owner pending", text: "Owner approved" }],
      },
      {
        outputSlide: 1,
        targetIndex: 2,
        operations: [{ type: "set-table-cell", row: 1, column: 1, expectedValue: "120", value: "135" }],
      },
      {
        outputSlide: 1,
        targetIndex: 3,
        operations: [
          { type: "set-chart-title", expectedTitle: "Revenue", title: "Revenue outlook" },
          { type: "set-chart-series-values", seriesIndex: 0, expectedValues: [120, 130], values: [135, 150] },
        ],
      },
      {
        outputSlide: 1,
        targetIndex: 4,
        operations: [{
          type: "replace-image",
          expectedSourceSha256: sha256(dataUrlBytes(ORIGINAL_PNG)),
          assetPath: path.basename(assetPath),
          assetSha256: sha256(replacementBytes),
        }],
      },
      { outputSlide: 2, targetIndex: 0, operations: [] },
    ],
  };
  await writeJson(planPath, plan);

  const result = await applyTemplateEditPlan({
    workspace: root,
    starterPath,
    manifestPath: starterManifestPath,
    planPath,
    out: outputPath,
    auditPath,
    previewDir,
    layoutDir,
    contactSheetPath,
    scale: 0.25,
  });
  assert.equal(result.audit.schema, "office-kit.template-edit-audit.v1");
  assert.equal(result.audit.status, "succeeded");
  assert.equal(result.audit.operation.count, 7);
  assert.equal(result.audit.operation.operations.every((operation) => operation.executed && operation.finalElementId), true);
  assert.deepEqual(result.audit.assets, [{
    path: assetPath,
    relativePath: "replacement.png",
    sha256: sha256(replacementBytes),
    bytes: replacementBytes.length,
  }]);
  assert.equal(result.audit.validation.allMappedTargetsCovered, true);
  assert.equal(result.audit.validation.finalExportReimported, true);
  assert.equal(result.audit.validation.immutableInputsRecheckedBeforePublication, true);
  assert.equal(result.audit.validation.untouchedSlideVisualsEquivalent, true);
  assert.deepEqual(result.audit.validation.untouchedSlides.map((entry) => entry.outputSlide), [2]);
  assert.deepEqual(await fs.readFile(sourcePath), sourceBytes);
  assert.deepEqual(await fs.readFile(starterPath), starterBytes);
  assert.deepEqual(await fs.readFile(starterManifestPath), starterManifestBytes);
  assert.deepEqual((await fs.readdir(previewDir)).sort(), ["final-slide-01.png", "final-slide-02.png"]);
  assert.deepEqual((await fs.readdir(layoutDir)).sort(), ["final-slide-01.layout.json", "final-slide-02.layout.json"]);
  assert.equal((await fs.readFile(contactSheetPath)).subarray(0, 8).toString("hex"), "89504e470d0a1a0a");
  assert.deepEqual(JSON.parse(await fs.readFile(auditPath, "utf8")), result.audit);

  const final = await PresentationFile.importPptx(await FileBlob.load(outputPath));
  const finalTitleId = result.audit.operation.operations.find((operation) => operation.type === "replace-text").finalElementId;
  const finalPositionId = result.audit.operation.operations.find((operation) => operation.type === "set-position").finalElementId;
  const finalOwnerId = result.audit.operation.operations.find((operation) => operation.type === "set-text").finalElementId;
  const finalTableId = result.audit.operation.operations.find((operation) => operation.type === "set-table-cell").finalElementId;
  const finalChartId = result.audit.operation.operations.find((operation) => operation.type === "set-chart-title").finalElementId;
  const finalImageId = result.audit.operation.operations.find((operation) => operation.type === "replace-image").finalElementId;
  assert.equal(final.resolve(finalTitleId).text.value, "Revenue outlook approved");
  assert.deepEqual(final.resolve(finalPositionId).position, { left: 48, top: 18, width: 270, height: 88 });
  assert.equal(final.resolve(finalOwnerId).text.value, "Owner approved");
  assert.equal(final.resolve(finalTableId).getCell(1, 1).value, "135");
  assert.equal(final.resolve(finalChartId).title, "Revenue outlook");
  assert.deepEqual(final.resolve(finalChartId).series[0].values, [135, 150]);
  assert.equal(sha256(dataUrlBytes(final.resolve(finalImageId).dataUrl)), sha256(replacementBytes));
  assert.equal(final.slides.items[1].shapes.items[0].text.value, "This slide must remain unchanged.");

  await assert.rejects(
    applyTemplateEditPlan({
      workspace: root,
      starterPath,
      manifestPath: starterManifestPath,
      planPath,
      out: outputPath,
      auditPath,
      previewDir,
      layoutDir,
      contactSheetPath,
    }),
    /already exists; refusing to overwrite/i,
  );

  const stalePlanPath = path.join(root, "stale-plan.json");
  const staleOutputPath = path.join(root, "stale-output.pptx");
  const staleAuditPath = path.join(root, "stale-output.audit.json");
  const stalePreviewDir = path.join(root, "stale-preview");
  const staleLayoutDir = path.join(root, "stale-layout");
  const stalePlan = structuredClone(plan);
  stalePlan.targets[0].operations[0].expectedText = "stale title";
  await writeJson(stalePlanPath, stalePlan);
  await assert.rejects(
    applyTemplateEditPlan({
      workspace: root,
      starterPath,
      manifestPath: starterManifestPath,
      planPath: stalePlanPath,
      out: staleOutputPath,
      auditPath: staleAuditPath,
      previewDir: stalePreviewDir,
      layoutDir: staleLayoutDir,
    }),
    /expectedText must occur exactly once/i,
  );
  for (const absent of [staleOutputPath, staleAuditPath, stalePreviewDir, staleLayoutDir]) {
    assert.equal(await fs.access(absent).then(() => true, () => false), false);
  }

  const deleteManifestPath = path.join(root, "delete.manifest.json");
  const deletePlanPath = path.join(root, "delete-plan.json");
  const deleteManifest = structuredClone(starterManifest);
  targetByIndex(deleteManifest, 1, 0).action = "delete";
  await writeJson(deleteManifestPath, deleteManifest);
  const deleteManifestBytes = await fs.readFile(deleteManifestPath);
  const deletePlan = structuredClone(plan);
  deletePlan.manifestSha256 = sha256(deleteManifestBytes);
  await writeJson(deletePlanPath, deletePlan);
  await assert.rejects(
    applyTemplateEditPlan({
      workspace: root,
      starterPath,
      manifestPath: deleteManifestPath,
      planPath: deletePlanPath,
      out: path.join(root, "delete-output.pptx"),
      auditPath: path.join(root, "delete-output.audit.json"),
      previewDir: path.join(root, "delete-preview"),
      layoutDir: path.join(root, "delete-layout"),
    }),
    /requests delete.*does not support/i,
  );
  assert.deepEqual(await fs.readFile(sourcePath), sourceBytes);
  assert.deepEqual(await fs.readFile(starterPath), starterBytes);
  assert.equal((await fs.readdir(root)).some((name) => name.startsWith(".office-kit-template-edit-")), false);
} finally {
  await fs.rm(root, { recursive: true, force: true });
}

console.log("presentation template edit transaction tests passed");
