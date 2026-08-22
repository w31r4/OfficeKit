import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

import { validateWindowsPptxLosslessEvidence } from "../scripts/validate-windows-pptx-lossless-evidence.mjs";

const manifest = JSON.parse(await readFile(path.resolve(import.meta.dirname, "../evals/pptx-lossless/manifest.v1.json"), "utf8"));
const collector = await readFile(path.resolve(import.meta.dirname, "../scripts/collect-windows-pptx-lossless-evidence.ps1"), "utf8");
assert.match(collector, /PowerPoint\.Application/u, "Windows collector must use the real PowerPoint COM host");
assert.match(collector, /SaveCopyAs/u, "Windows collector must save a distinct PowerPoint copy");
assert.match(collector, /Slide.*Export|\.Export\(/u, "Windows collector must export PowerPoint slide images");
assert.match(collector, /Get-FileHash/u, "Windows collector must hash rendered evidence");
assert.match(collector, /human-observed-windows-powerpoint/u, "Windows collector must emit the human-observed evidence method");
assert.match(collector, /UTF8Encoding[\s\S]*WriteAllText/u, "Windows collector must emit BOM-free JSON for the Node validator");
const checkedAt = "2026-08-22T12:00:00Z";
const commit = "0123456789abcdef0123456789abcdef01234567";
const sourceIds = ["suanzhi-future-2026", "blue-gray-acid-template", "mckinsey-customer-loyalty"];
const pageComparisons = sourceIds.flatMap((id, sourceIndex) => {
  const source = manifest.sources.find((candidate) => candidate.id === id);
  const targetPage = Number(String(source.targets[0].nodeId).match(/^presentation\/slide\/(\d+)/u)?.[1]);
  return Array.from({ length: source.inventory.slideCount }, (_, index) => {
    const page = index + 1;
    const target = targetPage === page;
    const pageHex = page.toString(16).padStart(2, "0");
    const sourcePixelSha256 = `${String.fromCharCode(97 + sourceIndex).repeat(62)}${pageHex}`;
    const outputPixelSha256 = target ? `${String.fromCharCode(100 + sourceIndex).repeat(62)}${pageHex}` : sourcePixelSha256;
    return { sourceId: id, page, target, pixelIdentical: !target, sourcePixelSha256, outputPixelSha256 };
  });
});
const checks = {
  opened: true,
  noRepairPrompt: true,
  browsedAllSlides: true,
  targetEditVisible: true,
  nonTargetPagesPixelIdentical: true,
  advancedObjectsPreserved: true,
  savedCopy: true,
  reopenedCopy: true,
  sourceProtected: true,
  unsupportedCapabilityFailClosed: true,
};
const evidence = {
  schema: "office-kit.windows-pptx-lossless-evidence.v1",
  method: "human-observed-windows-powerpoint",
  checkedAt,
  commit,
  host: {
    platform: "win32-x64",
    observedAt: checkedAt,
    powerpoint: { installed: true, version: "Microsoft PowerPoint 16.0 (Build 18025)" },
  },
  visualReview: {
    observedAt: checkedAt,
    renderer: "Microsoft PowerPoint",
    pagesCompared: 48,
    evidencePath: "evidence/windows-pptx-lossless/pages.json",
    pageComparisons,
  },
  sources: sourceIds.map((id, index) => {
    const source = manifest.sources.find((candidate) => candidate.id === id);
    return {
      id,
      sourceSha256: source.sha256,
      sourcePath: `C:\\OfficeKit\\assets\\${source.fileName}`,
      outputPath: `C:\\OfficeKit\\outputs\\${id}-edited.pptx`,
      target: {
        nodeId: source.targets[0].nodeId,
        operation: "native-leaf-edit",
      },
      checks: { ...checks },
      evidencePath: `evidence/windows-pptx-lossless/${index + 1}-${id}.json`,
    };
  }),
};

assert.deepEqual(validateWindowsPptxLosslessEvidence(evidence, { expectedCommit: commit, manifest }), {
  schema: "office-kit.windows-pptx-lossless-evidence.v1",
  checkedAt,
  commit,
  platform: "win32-x64",
  powerpointVersion: "Microsoft PowerPoint 16.0 (Build 18025)",
  sources: sourceIds,
  pagesCompared: 48,
});

for (const mutation of [
  (value) => { value.method = "mock"; },
  (value) => { value.host.platform = "darwin-arm64"; },
  (value) => { value.host.observedAt = "2026-08-23T12:00:00Z"; },
  (value) => { value.visualReview.renderer = "LibreOffice"; },
  (value) => { value.visualReview.pagesCompared = 47; },
  (value) => { value.visualReview.pageComparisons.pop(); },
  (value) => { value.visualReview.pageComparisons[0].outputPixelSha256 = value.visualReview.pageComparisons[0].sourcePixelSha256; },
  (value) => { value.visualReview.pageComparisons[1].pixelIdentical = false; },
  (value) => { value.sources[0].sourceSha256 = "0".repeat(64); },
  (value) => { value.sources[1].checks.nonTargetPagesPixelIdentical = false; },
  (value) => { value.sources.pop(); },
  (value) => { value.sources[0].sourcePath = "relative\\source.pptx"; },
  (value) => { value.sources[2].outputPath = value.sources[2].sourcePath; },
]) {
  const invalid = structuredClone(evidence);
  mutation(invalid);
  assert.throws(() => validateWindowsPptxLosslessEvidence(invalid, { expectedCommit: commit, manifest }), /evidence|Windows|PowerPoint|renderer|SHA|source|true|three|platform|output|pages|frozen|records|comparison|target/i);
}

assert.throws(
  () => validateWindowsPptxLosslessEvidence(evidence, { expectedCommit: "fedcba9876543210fedcba9876543210fedcba98", manifest }),
  /checked-out commit/,
);

console.log("Windows PPTX lossless evidence gate ok");
