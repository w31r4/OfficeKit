import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

import { validateWindowsPptxLosslessEvidence } from "../scripts/validate-windows-pptx-lossless-evidence.mjs";

const manifest = JSON.parse(await readFile(path.resolve(import.meta.dirname, "../evals/pptx-lossless/manifest.v1.json"), "utf8"));
const checkedAt = "2026-08-22T12:00:00Z";
const commit = "0123456789abcdef0123456789abcdef01234567";
const sourceIds = ["suanzhi-future-2026", "blue-gray-acid-template", "mckinsey-customer-loyalty"];
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
  (value) => { value.sources[0].sourceSha256 = "0".repeat(64); },
  (value) => { value.sources[1].checks.nonTargetPagesPixelIdentical = false; },
  (value) => { value.sources.pop(); },
  (value) => { value.sources[0].sourcePath = "relative\\source.pptx"; },
  (value) => { value.sources[2].outputPath = value.sources[2].sourcePath; },
]) {
  const invalid = structuredClone(evidence);
  mutation(invalid);
  assert.throws(() => validateWindowsPptxLosslessEvidence(invalid, { expectedCommit: commit, manifest }), /evidence|Windows|PowerPoint|renderer|SHA|source|true|three|platform|output|pages|frozen/i);
}

assert.throws(
  () => validateWindowsPptxLosslessEvidence(evidence, { expectedCommit: "fedcba9876543210fedcba9876543210fedcba98", manifest }),
  /checked-out commit/,
);

console.log("Windows PPTX lossless evidence gate ok");
