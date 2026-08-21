import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";

const manifestPath = path.resolve("evals/pptx-lossless/manifest.v1.json");
const manifestBytes = await readFile(manifestPath);
const manifest = JSON.parse(manifestBytes.toString("utf8"));
const evidence = JSON.parse(await readFile(path.resolve("evals/pptx-lossless/evidence.v1.json"), "utf8"));
const controls = JSON.parse(await readFile(path.resolve("evals/pptx-lossless/controls.v1.json"), "utf8"));
const designProfiles = JSON.parse(await readFile(path.resolve("evals/pptx-lossless/design-profiles.v1.json"), "utf8"));
const sourceReuse = JSON.parse(await readFile(path.resolve("evals/pptx-lossless/source-reuse.v1.json"), "utf8"));
const sourceComponentReuse = JSON.parse(await readFile(path.resolve("evals/pptx-lossless/source-component-reuse.v1.json"), "utf8"));
const svgText = JSON.parse(await readFile(path.resolve("evals/pptx-lossless/svg-text.v1.json"), "utf8"));

assert.equal(manifest.schema, "office-kit/pptx-lossless-benchmark/v1");
assert.equal(manifest.sources.length, 4);
assert.equal(new Set(manifest.sources.map((source) => source.id)).size, 4);
assert.equal(new Set(manifest.sources.map((source) => source.sha256)).size, 4);
assert.equal(manifest.sources.filter((source) => source.sourceKind === "external").length, 3);
assert.equal(manifest.sources.filter((source) => source.sourceKind === "repository-supplemental").length, 1);
assert.equal(manifest.sources.reduce((count, source) => count + source.targets.length, 0), 12);

for (const source of manifest.sources) {
  assert.match(source.id, /^[a-z0-9][a-z0-9-]+$/u);
  assert.equal(path.isAbsolute(source.fileName), false);
  assert.equal(source.fileName.includes(".."), false);
  assert.match(source.sha256, /^[a-f0-9]{64}$/u);
  assert.equal(Number.isSafeInteger(source.bytes) && source.bytes > 0, true);
  assert.equal(source.inventory.partCount, Object.keys(source.inventory.partHashes).length);
  assert.equal(Array.isArray(source.editableNodes), true);
  assert.equal("sourceBytes" in source, false);
  for (const [partPath, sha256] of Object.entries(source.inventory.partHashes)) {
    assert.equal(path.posix.isAbsolute(partPath), false);
    assert.equal(partPath.split("/").includes(".."), false);
    assert.match(sha256, /^[a-f0-9]{64}$/u);
  }
  for (const target of source.targets) {
    assert.match(target.nodeId, /^presentation\/slide\/[1-9][0-9]*\/element\/[1-9][0-9]*(?:\/.*)?$/u);
    if ((target.operation ?? "text") === "text") {
      assert.equal(typeof target.expected, "string");
      assert.equal(typeof target.value, "string");
      assert.equal(target.search === undefined || typeof target.search === "string", true);
      assert.equal(target.result === undefined || typeof target.result === "string", true);
      assert.notEqual(target.search ?? target.expected, target.value);
      assert.equal(source.editableNodes.some((node) => node.id === target.nodeId && node.text === target.expected), true);
    } else {
      assert.match(target.operation, /^native(?:Leaf|Leaves)$/u);
      const leaves = target.operation === "nativeLeaves" ? target.leaves : [target];
      assert.equal(Array.isArray(leaves) && leaves.length > 0 && leaves.length <= 8, true);
      assert.equal(new Set(leaves.map((leaf) => leaf.leafKind)).size, leaves.length);
      for (const leaf of leaves) {
        assert.match(leaf.leafKind, /^(text|chartTitleText|chartDataValue|diagramText|leftEmu|topEmu|widthEmu|heightEmu)$/u);
        if (leaf.leafKind === "text" || leaf.leafKind === "chartTitleText" || leaf.leafKind === "diagramText") {
          assert.equal(typeof leaf.expectedValue, "string");
          assert.equal(typeof leaf.value, "string");
        } else {
          assert.equal(Number.isSafeInteger(leaf.expectedValue), true);
          assert.equal(Number.isSafeInteger(leaf.value), true);
        }
        if (leaf.leafKind === "chartDataValue") {
          assert.equal(Number.isSafeInteger(leaf.seriesIndex) && leaf.seriesIndex >= 0, true);
          assert.equal(Number.isSafeInteger(leaf.pointIndex) && leaf.pointIndex >= 0, true);
        }
        if (leaf.leafKind === "diagramText") {
          assert.match(leaf.diagramNodeId, /^\{[A-F0-9-]{36}\}$/u);
          assert.equal(Number.isSafeInteger(leaf.runIndex) && leaf.runIndex >= 0, true);
        }
        assert.notEqual(leaf.expectedValue, leaf.value);
      }
    }
  }
}

assert.equal(
  manifest.sources.some((source) => source.targets.some((target) =>
    source.editableNodes.some((node) => node.id === target.nodeId && node.depth > 0))),
  true,
);

assert.equal(evidence.schema, "office-kit/pptx-lossless-evidence/v1");
assert.equal(evidence.manifestSha256, createHash("sha256").update(manifestBytes).digest("hex"));
assert.equal(evidence.repetitionsPerTarget, 3);
assert.deepEqual(Object.values(evidence.runnerContract), [true, true, true, true, true, true, true, true, true, true]);
assert.equal(evidence.sources.length, manifest.sources.length);
assert.equal(evidence.sources.reduce((count, source) => count + source.targets.length, 0), 12);
for (const source of evidence.sources) {
  const declared = manifest.sources.find((candidate) => candidate.id === source.id);
  assert.ok(declared);
  assert.equal(source.sourceSha256, declared.sha256);
  assert.equal(source.noOpByteIdentical, true);
  assert.equal(source.targets.length, declared.targets.length);
  for (const target of source.targets) {
    assert.equal(declared.targets.some((candidate) => candidate.id === target.id), true);
    assert.match(target.outputSha256, /^[a-f0-9]{64}$/u);
    assert.equal(Array.isArray(target.repetitionOutputSha256), true);
    assert.equal(target.repetitionOutputSha256.length, evidence.repetitionsPerTarget);
    assert.equal(new Set(target.repetitionOutputSha256).size, 1);
    assert.deepEqual(target.repetitionOutputSha256, [target.outputSha256, target.outputSha256, target.outputSha256]);
    assert.equal(Array.isArray(target.repetitionEditPlanSha256), true);
    assert.equal(target.repetitionEditPlanSha256.length, evidence.repetitionsPerTarget);
    assert.equal(new Set(target.repetitionEditPlanSha256).size, 1);
    if (target.leafKind === "chartDataValue") {
      assert.deepEqual(target.changedParts, ["ppt/charts/chart4.xml", "ppt/embeddings/Microsoft_Excel____1.xlsx"]);
    } else if (target.leafKind === "diagramText") {
      assert.deepEqual(target.changedParts, ["ppt/diagrams/strategy-data.xml"]);
    } else {
      assert.equal(target.changedParts.length, 1);
      assert.match(target.changedParts[0], /^ppt\/(?:slides\/slide[1-9][0-9]*|charts\/chart[1-9][0-9]*)\.xml$/u);
    }
  }
}

assert.equal(controls.schema, "office-kit/pptx-lossless-controls/v1");
assert.equal(controls.nativeRenderer.office, "LibreOffice");
assert.equal(controls.nativeRenderer.raster, "Poppler");
assert.equal(controls.nativeRenderer.dpi, 144);
assert.equal(controls.nativeVisualEvidence.length, manifest.sources.length);
for (const source of controls.nativeVisualEvidence) {
  const declared = manifest.sources.find((candidate) => candidate.id === source.id);
  assert.ok(declared);
  assert.equal(source.nonTargetPagesPixelIdentical, true);
  assert.equal(source.targets.length, declared.targets.length);
  for (const target of source.targets) {
    assert.equal(declared.targets.some((candidate) => candidate.id === target.id), true);
    assert.match(target.targetPageVisualState, /^(changed|unchanged-in-libreoffice)$/u);
    assert.equal(Number.isSafeInteger(target.differentPixels) && target.differentPixels >= 0, true);
  }
}
assert.equal(controls.historicalHtmlRebuildControl.status, "available-historical");
assert.match(controls.historicalHtmlRebuildControl.reportSha256, /^[a-f0-9]{64}$/u);
assert.equal(controls.kimiPptdControl.status, "not-available");
assert.equal(controls.kimiPptdControl.affectsOfficeKitAcceptance, false);
assert.equal(controls.quickLookTargetSpotCheck.status, "passed-limited");
assert.equal(controls.quickLookTargetSpotCheck.source, "suanzhi-future-2026");
assert.equal(controls.quickLookTargetSpotCheck.target, "title");
assert.match(controls.quickLookTargetSpotCheck.sourceThumbnailSha256, /^[a-f0-9]{64}$/u);
assert.match(controls.quickLookTargetSpotCheck.outputThumbnailSha256, /^[a-f0-9]{64}$/u);
assert.equal(controls.quickLookTargetSpotCheck.differentPixels > 0, true);

assert.equal(designProfiles.schema, "office-kit/pptx-design-profile/v1");
assert.equal(designProfiles.profiles.length, 3);
for (const profile of designProfiles.profiles) {
  const declared = manifest.sources.find((source) => source.id === profile.source.id);
  assert.ok(declared, `Missing benchmark manifest entry for design profile ${profile.source.id}`);
  assert.equal(declared.sha256, profile.source.sha256);
  assert.equal(declared.inventory.slideCount, profile.evidence.package.slideCount);
  assert.equal(profile.canvas.aspectRatio > 0, true);
  assert.equal(profile.designLanguage.density.slides, profile.evidence.package.slideCount);
  assert.equal(Array.isArray(profile.layoutFamilies), true);
  assert.equal(Array.isArray(profile.slideArchetypes), true);
  assert.equal(profile.slideArchetypes.length, profile.evidence.package.slideCount);
  assert.equal(profile.nativeOpaque.count >= 0, true);
  assert.equal(Object.keys(profile.evidence.structuralPartHashes).length > 0, true);
  assert.ok(profile.componentCandidates.total >= 0);
  assert.equal(profile.componentCandidates.total, Object.values(profile.componentCandidates.statuses).reduce((sum, count) => sum + count, 0));
  assert.ok(profile.componentCandidates.inspectOnlyCandidateIds.every((id) => /^pc_[0-9a-f]{32}$/u.test(id)));
}

assert.equal(sourceReuse.schema, "office-kit/pptx-source-reuse-evidence/v1");
assert.equal(sourceReuse.sources.length, 3);
for (const reuse of sourceReuse.sources) {
  const declared = manifest.sources.find((source) => source.id === reuse.id);
  assert.ok(declared, `Missing benchmark manifest entry for source reuse ${reuse.id}`);
  assert.equal(declared.sha256, reuse.sourceSha256);
  assert.match(reuse.sourceSlidePart, /^ppt\/slides\/slide[1-9][0-9]*\.xml$/u);
  assert.equal(reuse.status, reuse.expected);
  if (reuse.status === "passed") {
    assert.equal(reuse.sourceSlideUnchanged, true);
    assert.equal(reuse.outputSlideCount, reuse.sourceSlideCount + 1);
    assert.deepEqual(reuse.nonTopologyChangedParts, []);
    assert.deepEqual(reuse.topologyChangedParts, ["[Content_Types].xml", "ppt/_rels/presentation.xml.rels", "ppt/presentation.xml"]);
    assert.equal(reuse.addedParts.some((part) => /^ppt\/slides\/slide\d+\.xml$/u.test(part)), true);
  } else {
    assert.match(reuse.blockedReason, /shared|referenced|unsupported/i);
  }
}

assert.equal(sourceComponentReuse.schema, "office-kit/pptx-source-component-reuse-evidence/v1");
assert.deepEqual(sourceComponentReuse.sources.map((source) => source.id), [
  "suanzhi-future-2026",
  "blue-gray-acid-template",
  "mckinsey-customer-loyalty",
]);
assert.equal(sourceComponentReuse.sources.find((source) => source.id === "blue-gray-acid-template")?.status, "passed");
assert.equal(sourceComponentReuse.sources.find((source) => source.id === "mckinsey-customer-loyalty")?.status, "passed");
assert.equal(sourceComponentReuse.sources.find((source) => source.id === "suanzhi-future-2026")?.status, "passed");
for (const reuse of sourceComponentReuse.sources) {
  const declared = manifest.sources.find((source) => source.id === reuse.id);
  assert.ok(declared);
  assert.equal(reuse.sourceSha256, declared.sha256);
  assert.equal(reuse.sourceSlideCount, declared.inventory.slideCount);
  assert.equal(reuse.candidateCount >= reuse.inspectOnlyCandidateCount, true);
  assert.equal(reuse.preflightBlockedCandidateCount <= reuse.inspectOnlyCandidateCount, true);
  assert.equal(reuse.preflightBlockedReasons.length, reuse.preflightBlockedCandidateCount);
  if (reuse.status === "blocked") {
    assert.equal(reuse.inspectOnlyCandidateCount > 0, true);
    assert.equal(reuse.preflightBlockedCandidateCount, reuse.inspectOnlyCandidateCount);
    assert.equal(reuse.failures.unsupported_presentation_component_reuse || 0, 0);
    assert.match(reuse.blockedReason, /preflight/i);
  } else {
    assert.equal(reuse.status, "passed");
    assert.match(reuse.candidateId, /^pc_[0-9a-f]{32}$/u);
    assert.equal(reuse.nonTargetPartMismatches.length, 0);
    assert.equal(reuse.reopenedSlideCount, reuse.sourceSlideCount + 1);
    assert.equal(reuse.cloneElementCount >= 1, true);
    assert.match(reuse.outputPackageContentSha256, /^[a-f0-9]{64}$/u);
    assert.match(reuse.outputSha256, /^[a-f0-9]{64}$/u);
  }
}

assert.equal(svgText.schema, "office-kit/pptx-svg-text-evidence/v1");
assert.deepEqual(svgText.sources.map((source) => source.id), [
  "suanzhi-future-2026",
  "blue-gray-acid-template",
  "mckinsey-customer-loyalty",
]);
for (const source of svgText.sources) {
  const declared = manifest.sources.find((candidate) => candidate.id === source.id);
  assert.ok(declared);
  assert.equal(source.sourceSha256, declared.sha256);
  assert.equal(source.sourceSlideCount, declared.inventory.slideCount);
  assert.equal(Number.isSafeInteger(source.svgTextNodeCount) && source.svgTextNodeCount >= 0, true);
  if (source.id === "mckinsey-customer-loyalty") {
    assert.equal(source.status, "passed");
    assert.equal(source.svgImageCount, 8);
    assert.equal(source.svgTextNodeCount, 250);
    assert.equal(source.reimported, true);
    assert.deepEqual(source.changedExistingParts, ["ppt/slides/_rels/slide1.xml.rels", "ppt/slides/slide1.xml"]);
    assert.equal(source.addedParts.length, 1);
    assert.match(source.addedParts[0], /^ppt\/media\/image\d+\.svg$/u);
  } else {
    assert.equal(source.status, "not-applicable");
    assert.equal(source.svgTextNodeCount, 0);
  }
}

console.log("PPTX lossless benchmark manifest smoke ok");
