import assert from "node:assert/strict";
import path from "node:path";
import { readFile } from "node:fs/promises";
import JSZip from "jszip";

import { buildPptxDesignProfile } from "../scripts/pptx-design-profile.mjs";
import { FileBlob, Presentation, PresentationFile } from "../src/index.mjs";

const fixture = path.resolve(import.meta.dirname, "../evals/assets/presentations/strategy-review.pptx");
const first = await buildPptxDesignProfile(fixture, { id: "smartart-canary" });
const second = await buildPptxDesignProfile(fixture, { id: "smartart-canary" });

assert.equal(first.schema, "office-kit/pptx-design-profile/v1");
assert.deepEqual(first, second, "design profiles must be deterministic");
assert.deepEqual(first.source, {
  id: "smartart-canary",
  fileName: "strategy-review.pptx",
  bytes: first.source.bytes,
  sha256: first.source.sha256,
});
assert.match(first.source.sha256, /^[a-f0-9]{64}$/u);
assert.equal(first.canvas.width, 1280);
assert.equal(first.canvas.height, 720);
assert.equal(first.evidence.package.slideCount, 4);
assert.equal(first.evidence.package.diagramCount >= 1, true);
assert.equal(first.nativeOpaque.kinds.diagram >= 1, true);
assert.equal(first.slideArchetypes.length, 4);
assert.equal(first.layoutFamilies.length >= 1, true);
assert.equal(Array.isArray(first.designLanguage.palette.scheme), true);
assert.equal(Array.isArray(first.designLanguage.palette.direct), true);
assert.equal(first.designLanguage.typography.fontSizePt.samples > 0, true);
assert.equal(first.designLanguage.density.slides, 4);
assert.equal(first.designLanguage.rhythm.normalizedUnits, "slide fraction rounded to 0.001");
assert.equal(first.reusableComponents.every((component) => component.count >= 2), true);
assert.equal(JSON.stringify(first).includes(fixture), false, "profile must not leak an absolute source path");

// A real imported slide may omit p:cSld/@name. The codec must preserve that
// absence instead of inventing "Slide N", otherwise a safe source-derived
// clone is rejected as a metadata mutation.
const sourceBytes = await readFile(fixture);
const sourceZip = await JSZip.loadAsync(sourceBytes);
const slideXml = await sourceZip.file("ppt/slides/slide1.xml").async("text");
const unnamedSlideXml = slideXml.replace(/(<p:cSld\b[^>]*?)\sname="[^"]*"/u, "$1");
assert.notEqual(unnamedSlideXml, slideXml);
sourceZip.file("ppt/slides/slide1.xml", unnamedSlideXml);
const unnamedBytes = await sourceZip.generateAsync({ type: "uint8array", compression: "STORE" });
const importedUnnamed = await PresentationFile.importPptx(new FileBlob(unnamedBytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }));
assert.equal(importedUnnamed.slides.items[0].name, "");
assert.equal(importedUnnamed.slides.items[0].cloneCapability.supported, true);
importedUnnamed.slides.items[0].duplicate();
const clonedUnnamed = await PresentationFile.exportPptx(importedUnnamed);
const clonedZip = await JSZip.loadAsync(clonedUnnamed.bytes);
assert.equal(await clonedZip.file("ppt/slides/slide1.xml").async("text"), unnamedSlideXml);
assert.equal(importedUnnamed.slides.count, 5);

// Imported design evidence exposes repeated visual primitives as stable,
// source-bound references. The v1 surface is deliberately inspect-only: it
// never turns a geometry match into permission to synthesize a partial graph.
const candidatePresentation = await PresentationFile.importPptx(new FileBlob(sourceBytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }));
const componentRecords = candidatePresentation.inspect({ includeComponentCandidates: true, maxChars: Infinity }).ndjson
  .split("\n").filter(Boolean).map((line) => JSON.parse(line)).filter((record) => record.kind === "componentCandidate");
assert.equal(componentRecords.length > 0, true);
assert.equal(componentRecords.every((record) => /^pc_[0-9a-f]{32}$/u.test(record.candidateId)), true);
assert.equal(componentRecords.every((record) => /^[0-9a-f]{64}$/u.test(record.sourceRevisionSha256)), true);
assert.equal(componentRecords.every((record) => record.mutationCapability?.supported === false), true);
assert.equal(componentRecords.every((record) => !JSON.stringify(record).includes("rawXml") && !JSON.stringify(record).includes("<p:")), true);
const safeCandidate = componentRecords.find((record) => record.status === "inspect-only");
assert.ok(safeCandidate, "fixture should expose at least one unambiguous inspect-only component candidate");
assert.deepEqual(candidatePresentation.resolveComponentCandidate(safeCandidate.candidateId), safeCandidate);
assert.equal(candidatePresentation.resolveComponentCandidate("pc_ffffffffffffffffffffffffffffffff"), undefined);
const noOpCandidateExport = await PresentationFile.exportPptx(candidatePresentation);
assert.deepEqual([...noOpCandidateExport.bytes], [...sourceBytes], "candidate inspection must not change source bytes");

// A source-bound reuse request carries the exact inspected revision and
// ownership evidence instead of allowing an Agent to clone by visual guess or
// array index. The pending clone remains a complete graph until export.
const reusePresentation = await PresentationFile.importPptx(new FileBlob(sourceBytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }));
const reusableSlide = reusePresentation.slides.items[0];
const reuseCapability = reusableSlide.cloneCapability;
assert.match(reuseCapability.sourceRevisionSha256, /^[0-9a-f]{64}$/u);
const reusedSlide = reusePresentation.reuseSourceSlide({
  slideId: reusableSlide.id,
  sourceRevisionSha256: reuseCapability.sourceRevisionSha256,
  expectedCloneCapability: reuseCapability,
});
assert.equal(reusedSlide.index, 1);
assert.equal(reusePresentation.slides.count, 5);
const reusedExport = await PresentationFile.exportPptx(reusePresentation);
const reusedZip = await JSZip.loadAsync(reusedExport.bytes);
assert.equal(
  await reusedZip.file("ppt/slides/slide1.xml").async("text"),
  slideXml,
  "source-bound slide reuse must preserve the original SlidePart bytes",
);
assert.equal((await PresentationFile.importPptx(reusedExport)).slides.count, 5);
const staleReusePresentation = await PresentationFile.importPptx(new FileBlob(sourceBytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }));
assert.throws(
  () => staleReusePresentation.reuseSourceSlide({ slideId: staleReusePresentation.slides.items[0].id, sourceRevisionSha256: "0".repeat(64) }),
  (error) => error?.code === "stale_presentation_source_revision",
);
const sourceFreeCandidateError = (() => {
  try {
    Presentation.create().inspect({ includeComponentCandidates: true });
    return undefined;
  } catch (error) {
    return error;
  }
})();
assert.equal(sourceFreeCandidateError?.code, "presentation_component_source_required");

console.log("PPTX design profile smoke ok");
