import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  DocumentFile,
  DocumentModel,
  FileBlob,
  PdfArtifact,
  reviewArtifact,
  Workbook,
} from "../src/index.mjs";
import { normalizePresentationAuthoringPlan } from "../src/cli/authoring-plan.mjs";
import { Presentation, PresentationFile } from "../src/presentation/index.mjs";

function authoringPlan(pageCount = 3, overrides = {}) {
  return {
    schema: "office-kit/presentation-authoring-plan/v1",
    mode: "create",
    brief: {
      audience: "Decision makers",
      purpose: "Choose a migration path",
      primaryJob: "decide",
      supportingJobs: ["align"],
      expectedOutcome: "Decision makers approve one bounded migration path",
      mediumFit: "strong",
      afterUse: "Decision record and delivery handoff",
      deliveryMode: "hybrid",
    },
    narrative: { thesis: "Use the bounded path" },
    design: {
      sourceMode: "self-directed",
      mechanismPacks: ["technical-architecture"],
      scenario: { primary: "technical-engineering", secondary: "analysis-decision" },
      direction: {
        name: "Traceable systems decision",
        rationale: "The audience needs architecture evidence tied directly to risk and ownership",
      },
      designGrammar: {
        palette: { strict: false, roles: { background: "#F8FAFC", accent: "#2563EB" } },
        typography: { strict: false, roles: { title: "Aptos Display", body: "Aptos" } },
        densityRhythm: "evidence then conclusion",
      },
    },
    pages: Array.from({ length: pageCount }, (_, index) => ({
      id: `page-${index + 1}`,
      readerTask: `Read page ${index + 1}`,
      claim: `Claim ${index + 1}`,
      evidence: [`Evidence ${index + 1}`],
      compositionIntent: index % 2 ? "split comparison" : "single focal statement",
      contentBudget: { maxCharacters: 2_000, maxObjects: 50 },
    })),
    editorial: { voice: "direct and evidence-led" },
    artifactRefs: [],
    recipe: "tasks/create.md",
    unresolved: [],
    nextAction: "Review the working draft",
    ...overrides,
  };
}

const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "officekit-review-"));
try {
  const document = DocumentModel.create({ paragraphs: ["Quarterly review document"] });

  const workbook = Workbook.create();
  workbook.worksheets.add("Data").getRange("A1:B2").values = [
    ["Metric", "Value"],
    ["Revenue", 42],
  ];

  const presentation = Presentation.create();
  presentation.slides.add({ name: "Review" }).shapes.add({
    geometry: "textbox",
    text: "Quarterly review presentation",
    position: { left: 20, top: 20, width: 400, height: 80 },
  });

  const pdf = PdfArtifact.create({ pages: [{ text: "Quarterly review PDF" }] });

  for (const [format, artifact, expectedText] of [
    ["docx", document, "Quarterly review document"],
    ["xlsx", workbook, "Revenue"],
    ["pptx", presentation, "Quarterly review presentation"],
    ["pdf", pdf, "Quarterly review PDF"],
  ]) {
    const report = await reviewArtifact(artifact, {
      outputPath: path.join(temporary, `review.${format}`),
      contentView: "anydoc",
      visualReview: "complete",
    });
    assert.equal(report.schemaVersion, 1);
    assert.equal(report.format, format);
    assert.equal(report.verdict, "passed", JSON.stringify(report, null, 2));
    assert.equal(report.semantic.status, "passed");
    assert.equal(report.structural.status, "passed");
    assert.equal(report.layout.status, "passed");
    assert.equal(report.contentView.status, "ready");
    assert.equal(report.contentView.provider, "anydoc");
    assert.equal(report.contentView.providerVersion, "0.1.3");
    assert.equal(report.contentView.sourceSha256, report.delivery.sha256);
    assert.match(report.contentView.markdown, new RegExp(expectedText, "i"));
    for (const step of ["6. Semantic", "7. Structural", "8. Layout", "9. Text reading", "10. Visual", "11. Delivery"]) {
      assert.match(report.summary.markdown, new RegExp(step.replace(".", "\\.")));
    }
  }

  const unavailableVisualReview = await reviewArtifact(document, {
    outputPath: path.join(temporary, "unavailable-visual.docx"),
  });
  assert.equal(unavailableVisualReview.verdict, "passed-with-limitations");
  assert.equal(unavailableVisualReview.visualReview, "unavailable");

  // Imported source-bound work often contains pre-existing geometry warnings.
  // A baseline may downgrade only the exact same issues; a newly introduced
  // issue must remain a hard review failure.
  const importedBaselineModel = Presentation.create();
  const importedBaselineSlide = importedBaselineModel.slides.add({ name: "Existing layout" });
  importedBaselineSlide.shapes.add({ geometry: "rect", text: "Existing background shape", position: { left: 200, top: 20, width: 220, height: 80 } });
  importedBaselineSlide.shapes.add({ geometry: "textbox", text: "Edit me", position: { left: 200, top: 20, width: 220, height: 80 } });
  const importedBaseline = await PresentationFile.exportPptx(importedBaselineModel);
  const importedEditedModel = await PresentationFile.importPptx(importedBaseline);
  importedEditedModel.slides.items[0].shapes.items[1].text.set("Edited without rebuilding");
  const importedEdited = await PresentationFile.exportPptx(importedEditedModel);
  const baselineReview = await reviewArtifact(new FileBlob(importedBaseline.bytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }), {
    format: "pptx",
    outputPath: path.join(temporary, "baseline.pptx"),
    layout: false,
    visualReview: "unavailable",
  });
  assert.equal(baselineReview.verdict, "failed");
  const baselineCompared = await reviewArtifact(new FileBlob(importedEdited.bytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }), {
    format: "pptx",
    outputPath: path.join(temporary, "baseline-compared.pptx"),
    baseline: new FileBlob(importedBaseline.bytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }),
    layout: false,
    visualReview: "unavailable",
  });
  assert.notEqual(baselineCompared.verdict, "failed", JSON.stringify(baselineCompared, null, 2));
  assert.ok(baselineCompared.baseline.matchedIssues > 0);
  assert.equal(baselineCompared.baseline.newIssues, 0);
  assert.ok(baselineCompared.semantic.issues.some((issue) => issue.preexisting === true));

  const sourceBoundCompared = await reviewArtifact(new FileBlob(importedEdited.bytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }), {
    format: "pptx",
    outputPath: path.join(temporary, "source-bound-compared.pptx"),
    source: new FileBlob(importedBaseline.bytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }),
    layout: false,
    visualReview: "unavailable",
  });
  assert.notEqual(sourceBoundCompared.verdict, "failed", JSON.stringify(sourceBoundCompared, null, 2));
  assert.ok(sourceBoundCompared.baseline.matchedIssues > 0);
  assert.equal(sourceBoundCompared.baseline.newIssues, 0);

  const introducedIssueModel = await PresentationFile.importPptx(importedBaseline);
  introducedIssueModel.slides.items[0].shapes.items[1].position = { ...introducedIssueModel.slides.items[0].shapes.items[1].position, width: 1_400 };
  const introducedIssue = await PresentationFile.exportPptx(introducedIssueModel);
  const newIssueReview = await reviewArtifact(new FileBlob(introducedIssue.bytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }), {
    format: "pptx",
    outputPath: path.join(temporary, "baseline-new-issue.pptx"),
    baseline: new FileBlob(importedBaseline.bytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }),
    layout: false,
    visualReview: "unavailable",
  });
  assert.equal(newIssueReview.verdict, "failed");
  assert.ok(newIssueReview.baseline.newIssues > 0);

  const plannedModel = Presentation.create();
  for (let index = 0; index < 3; index += 1) {
    const slide = plannedModel.slides.add({ name: `Planned ${index + 1}` });
    slide.shapes.add({
      geometry: index === 1 ? "ellipse" : "textbox",
      text: `Decision evidence ${index + 1}`,
      position: { left: 40 + index * 20, top: 40, width: 320 + index * 30, height: 90 },
      fill: index === 1 ? "#DBEAFE" : "#FFFFFF",
      textStyle: { fontFamily: "Aptos", color: "#17202A" },
    });
  }
  const plan = authoringPlan();
  const plannedReview = await reviewArtifact(plannedModel, {
    authoringPlan: plan,
    outputPath: path.join(temporary, "planned.pptx"),
    layout: false,
    visualReview: "unavailable",
  });
  assert.notEqual(plannedReview.verdict, "failed", JSON.stringify(plannedReview, null, 2));
  assert.equal(plannedReview.design.planSha256, normalizePresentationAuthoringPlan(plan).sha256);
  assert.equal(plannedReview.design.pageSignatures.length, 3);
  assert.equal(plannedReview.design.changedPageIds.length, 0);
  assert.equal(plannedReview.design.strategy.primaryJob, "decide");
  assert.equal(plannedReview.design.strategy.scenario.primary, "technical-engineering");
  assert.equal(plannedReview.design.strategy.direction.name, "Traceable systems decision");
  assert.equal(plannedReview.design.layers.communication.status, "passed");
  assert.equal(plannedReview.design.layers.narrative.status, "passed");
  assert.match(plannedReview.summary.markdown, /Authoring-plan design checks/u);
  assert.match(plannedReview.summary.markdown, /Strategy: decide; scenario: technical-engineering; direction: Traceable systems decision/u);

  const motionReviewModel = Presentation.create();
  const motionReviewSlide = motionReviewModel.slides.add({ name: "Motion review" });
  const motionReviewTarget = motionReviewSlide.shapes.add({
    geometry: "rect",
    name: "Primary risk",
    text: "Liquidity risk",
    position: { left: 120, top: 100, width: 420, height: 180 },
  });
  motionReviewSlide.animations.add(motionReviewTarget, { effect: "pulse", phase: "emphasis", start: "afterPrevious" });
  const motionPlan = authoringPlan(1);
  motionPlan.brief.deliveryMode = "live";
  motionPlan.design.motionPolicy = "adaptive";
  motionPlan.pages[0].compositionIntent = "native vector focus with one dominant risk statement";
  motionPlan.pages[0].motionIntent = {
    purpose: "focus",
    recipe: "focus-pulse",
    transition: "none",
    units: [{ id: "risk-pulse", targetRole: "primary risk", order: 1, start: "afterPrevious" }],
  };
  const motionReview = await reviewArtifact(motionReviewModel, {
    authoringPlan: motionPlan,
    outputPath: path.join(temporary, "motion-review.pptx"),
    layout: false,
    playbackEvidence: "structural",
    visualReview: "unavailable",
  });
  assert.notEqual(motionReview.verdict, "failed", JSON.stringify(motionReview, null, 2));
  assert.equal(motionReview.motion.animationCount, 1);
  assert.equal(motionReview.playbackEvidence, "structural");
  assert.match(motionReview.summary.markdown, /Motion and playback checks/u);

  const readerMotionPlan = structuredClone(motionPlan);
  readerMotionPlan.brief.deliveryMode = "reader";
  const readerMotionReview = await reviewArtifact(motionReviewModel, {
    authoringPlan: readerMotionPlan,
    outputPath: path.join(temporary, "reader-motion-review.pptx"),
    layout: false,
    visualReview: "unavailable",
  });
  assert.equal(readerMotionReview.verdict, "failed");
  assert.ok(readerMotionReview.motion.issues.some((issue) => issue.type === "readerMotionUnauthorized"));

  const typographyFloorModel = Presentation.create();
  const typographyFloorSlide = typographyFloorModel.slides.add({ name: "Typography floor" });
  typographyFloorSlide.shapes.add({
    geometry: "textbox",
    name: "Source note",
    text: "Supporting note",
    position: { left: 40, top: 40, width: 320, height: 80 },
    textStyle: { fontFamily: "Aptos", fontSize: 14, color: "#17202A" },
  });
  typographyFloorSlide.shapes.add({
    geometry: "textbox",
    name: "Body copy",
    text: "Readable body",
    position: { left: 40, top: 140, width: 320, height: 80 },
    textStyle: { fontFamily: "Aptos", fontSize: 19, color: "#17202A" },
  });
  const typographyFloorPlan = authoringPlan();
  typographyFloorPlan.pages = [typographyFloorPlan.pages[0]];
  typographyFloorPlan.design.designGrammar.typography = {
    strict: false,
    roles: { title: "Aptos Display", body: "Aptos" },
    minimumBodyFontSize: 20,
    minimumCaptionFontSize: 18,
  };
  const typographyFloorReview = await reviewArtifact(typographyFloorModel, {
    authoringPlan: typographyFloorPlan,
    outputPath: path.join(temporary, "typography-floor.pptx"),
    layout: false,
    visualReview: "unavailable",
  });
  assert.equal(typographyFloorReview.verdict, "failed");
  assert.ok(typographyFloorReview.design.issues.some((issue) => issue.type === "minimumFontSize" && issue.actual === 14 && issue.required === 18 && issue.role === "caption"));
  assert.ok(typographyFloorReview.design.issues.some((issue) => issue.type === "minimumFontSize" && issue.actual === 19 && issue.required === 20 && issue.role === "body"));

  const captionFloorModel = Presentation.create();
  captionFloorModel.slides.add({ name: "Caption floor" }).shapes.add({
    geometry: "textbox",
    name: "Source note",
    text: "Readable source note",
    position: { left: 40, top: 40, width: 320, height: 80 },
    textStyle: { fontFamily: "Aptos", fontSize: 19, color: "#17202A" },
  });
  const captionFloorPlan = authoringPlan(1);
  captionFloorPlan.design.designGrammar.typography = {
    strict: false,
    roles: { title: "Aptos Display", body: "Aptos" },
    minimumBodyFontSize: 22,
    minimumCaptionFontSize: 18,
  };
  const captionFloorReview = await reviewArtifact(captionFloorModel, {
    authoringPlan: captionFloorPlan,
    outputPath: path.join(temporary, "caption-floor.pptx"),
    layout: false,
    visualReview: "unavailable",
  });
  assert.notEqual(captionFloorReview.verdict, "failed", JSON.stringify(captionFloorReview, null, 2));
  assert.equal(captionFloorReview.design.issues.some((issue) => issue.type === "minimumFontSize"), false);

  const warningModel = Presentation.create();
  for (let slideIndex = 0; slideIndex < 3; slideIndex += 1) {
    const slide = warningModel.slides.add({ name: `Warning ${slideIndex + 1}` });
    const count = slideIndex === 1 ? 12 : 1;
    for (let item = 0; item < count; item += 1) {
      slide.shapes.add({
        geometry: "rect",
        text: item === 0 ? `Metric card ${slideIndex + 1}` : `Detail ${item + 1}`,
        position: { left: 20 + (item % 3) * 170, top: 20 + Math.floor(item / 3) * 100, width: 150, height: 80 },
      });
    }
    slide.shapes.add({
      geometry: "ellipse",
      name: `Dominant orbit ${slideIndex + 1}`,
      position: { left: 660, top: 180, width: 600, height: 400 },
      fill: "#E2E8F0",
      line: { color: "#64748B", width: 1 },
    });
    if (slideIndex === 0) {
      slide.shapes.add({
        geometry: "rect",
        name: "Large empty boundary",
        position: { left: 20, top: 250, width: 600, height: 390 },
        fill: null,
        line: { color: "#334155", width: 2 },
      });
    }
    if (slideIndex === 2) {
      slide.shapes.add({
        geometry: "rect",
        name: "Undifferentiated report container",
        text: "Quarterly risk\nOwner and next action",
        position: { left: 20, top: 260, width: 550, height: 320 },
        fill: "#FFFFFF",
        line: { color: "#CBD5E1", width: 1 },
        textStyle: { fontFamily: "Aptos", fontSize: 20 },
      });
    }
  }
  const warningReview = await reviewArtifact(warningModel, {
    authoringPlan: authoringPlan(),
    outputPath: path.join(temporary, "warning-signals.pptx"),
    layout: false,
    visualReview: "unavailable",
  });
  assert.notEqual(warningReview.verdict, "failed", JSON.stringify(warningReview, null, 2));
  assert.ok(warningReview.design.issues.some((issue) => issue.type === "densityRhythmJump" && issue.severity === "warning"));
  assert.ok(warningReview.design.issues.some((issue) => issue.type === "cardWallPattern" && issue.severity === "warning"));
  assert.ok(warningReview.design.issues.some((issue) => issue.type === "repeatedDominantGeometry" && issue.severity === "warning"));
  assert.ok(warningReview.design.issues.some((issue) => issue.type === "largeHollowContainer" && issue.severity === "warning"));
  assert.ok(warningReview.design.issues.some((issue) => issue.type === "weakTextContainerHierarchy" && issue.severity === "warning"));
  assert.ok(warningReview.design.issues.some((issue) => issue.type === "repeatedTitleForm" && issue.severity === "warning"));
  assert.equal(warningReview.design.layers.visual.status, "passed-with-warnings");

  const wrongPageCount = await reviewArtifact(plannedModel, {
    authoringPlan: authoringPlan(2),
    outputPath: path.join(temporary, "wrong-page-count.pptx"),
    layout: false,
    visualReview: "unavailable",
  });
  assert.equal(wrongPageCount.verdict, "failed");
  assert.ok(wrongPageCount.design.issues.some((issue) => issue.type === "authoringPlanPageCount"));

  const unresolvedPlan = authoringPlan();
  unresolvedPlan.unresolved = [{ id: "missing-evidence", required: true }];
  const unresolvedReview = await reviewArtifact(plannedModel, {
    authoringPlan: unresolvedPlan,
    outputPath: path.join(temporary, "unresolved.pptx"),
    layout: false,
    visualReview: "unavailable",
  });
  assert.equal(unresolvedReview.verdict, "failed");
  assert.ok(unresolvedReview.design.issues.some((issue) => issue.type === "requiredAuthoringDecision"));

  const budgetPlan = authoringPlan();
  budgetPlan.pages[0].contentBudget.maxCharacters = 1;
  const budgetReview = await reviewArtifact(plannedModel, {
    authoringPlan: budgetPlan,
    outputPath: path.join(temporary, "budget.pptx"),
    layout: false,
    visualReview: "unavailable",
  });
  assert.equal(budgetReview.verdict, "failed");
  assert.ok(budgetReview.design.issues.some((issue) => issue.type === "contentBudgetCharacters"));

  const strictPlan = authoringPlan();
  strictPlan.design.designGrammar.palette = { strict: true, allowedColors: ["#000000"] };
  strictPlan.design.designGrammar.typography = { strict: true, allowedFonts: ["Courier New"] };
  const strictReview = await reviewArtifact(plannedModel, {
    authoringPlan: strictPlan,
    outputPath: path.join(temporary, "strict.pptx"),
    layout: false,
    visualReview: "unavailable",
  });
  assert.equal(strictReview.verdict, "failed");
  assert.ok(strictReview.design.issues.some((issue) => issue.type === "strictPaletteViolation" || issue.type === "strictTypographyViolation"));

  const plannedBaseline = await PresentationFile.exportPptx(plannedModel);
  const locallyEdited = await PresentationFile.importPptx(plannedBaseline);
  locallyEdited.slides.items[0].shapes.items[0].text.set("Sharpened decision evidence");
  const locallyEditedFile = await PresentationFile.exportPptx(locallyEdited);
  const localReview = await reviewArtifact(locallyEditedFile, {
    authoringPlan: plan,
    changedPageIds: ["page-1"],
    baseline: plannedBaseline,
    outputPath: path.join(temporary, "local-edit.pptx"),
    layout: false,
    visualReview: "unavailable",
  });
  assert.notEqual(localReview.verdict, "failed", JSON.stringify(localReview, null, 2));
  assert.equal(localReview.design.issues.some((issue) => issue.type === "undeclaredPageChange"), false);

  const outOfScope = await PresentationFile.importPptx(locallyEditedFile);
  outOfScope.slides.items[1].shapes.items[0].text.set("Undeclared second-page edit");
  const outOfScopeFile = await PresentationFile.exportPptx(outOfScope);
  const outOfScopeReview = await reviewArtifact(outOfScopeFile, {
    authoringPlan: plan,
    changedPageIds: ["page-1"],
    baseline: plannedBaseline,
    outputPath: path.join(temporary, "out-of-scope.pptx"),
    layout: false,
    visualReview: "unavailable",
  });
  assert.equal(outOfScopeReview.verdict, "failed");
  assert.ok(outOfScopeReview.design.issues.some((issue) => issue.type === "undeclaredPageChange" && issue.pageId === "page-2"));

  await assert.rejects(
    reviewArtifact(document, { authoringPlan: plan }),
    /available only for Presentation review/u,
  );

  const disabledContentView = await reviewArtifact(document, {
    outputPath: path.join(temporary, "disabled-content.docx"),
    contentView: false,
    visualReview: "complete",
  });
  assert.equal(disabledContentView.contentView.status, "not-requested");
  assert.equal(disabledContentView.verdict, "passed");

  const longDocument = DocumentModel.create({ paragraphs: [`Review ${"content ".repeat(500)}`] });
  const truncated = await reviewArtifact(longDocument, {
    outputPath: path.join(temporary, "truncated.docx"),
    contentView: "anydoc",
    maxContentChars: 240,
    maxSummaryChars: 3_000,
    visualReview: "complete",
  });
  assert.equal(truncated.contentView.truncated, true);
  assert.ok(truncated.contentView.originalChars > truncated.contentView.chars);
  assert.match(truncated.contentView.markdown, /content view truncated/);
  assert.match(truncated.summary.markdown, /11\. Delivery review/);

  const savedPath = path.join(temporary, "source.docx");
  await (await DocumentFile.exportDocx(document)).save(savedPath);
  const collision = await reviewArtifact(savedPath, {
    source: savedPath,
    visualReview: "complete",
  });
  assert.equal(collision.delivery.ok, false);
  assert.equal(collision.verdict, "failed");
  assert.ok(collision.delivery.issues.some((issue) => issue.type === "inputOutputCollision"));

  await assert.rejects(
    reviewArtifact(new Uint8Array(20), { format: "pdf", maxBytes: 10 }),
    /exceeds maxBytes/,
  );
  await assert.rejects(
    reviewArtifact(document, { maxBytes: Number.NaN }),
    /maxBytes must be a positive safe integer/,
  );
  await assert.rejects(
    reviewArtifact(document, { visualReview: "ocr-complete" }),
    /visualReview must be complete, unavailable, or requires-human/,
  );

  const corrupt = await reviewArtifact(
    new Blob(["not an OOXML ZIP"], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }),
    { outputPath: path.join(temporary, "corrupt.docx"), contentView: "anydoc", visualReview: "complete" },
  );
  assert.equal(corrupt.verdict, "failed");
  assert.equal(corrupt.structural.status, "failed");
  assert.equal(corrupt.contentView.status, "blocked");

  const rootProbe = spawnSync(process.execPath, [
    "--input-type=module",
    "-e",
    "import Module from 'node:module'; await import('./src/index.mjs'); console.log(JSON.stringify(Object.keys(Module._cache).filter((name) => /firecrawl|anydoc/i.test(name))))",
  ], {
    cwd: path.resolve(import.meta.dirname, ".."),
    encoding: "utf8",
    env: process.env,
  });
  assert.equal(rootProbe.status, 0, rootProbe.stderr);
  assert.deepEqual(JSON.parse(rootProbe.stdout), [], "root import must not initialize AnyDoc or its native binding");
} finally {
  await fs.rm(temporary, { recursive: true, force: true });
}

console.log("post-edit review smoke ok");
