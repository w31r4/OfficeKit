import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

import { FileBlob, Presentation, PresentationFile } from "../src/index.mjs";

const ONE_PIXEL_PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";

const authored = Presentation.create({
  slideSize: { width: 1280, height: 720 },
  theme: { colors: { accent1: "#2563eb" }, fonts: { major: "Arial", minor: "Arial" } },
});
const firstSlide = authored.slides.add({ name: "Profile source" });
for (const [text, top] of [["Quarterly review", 36], ["Supporting detail", 36]]) {
  firstSlide.shapes.add({
    name: "Title",
    position: { left: 48, top, width: 480, height: 64 },
    text: [{ runs: [{ text, style: { fontFamily: "Arial", fontSize: 28, color: "#2563eb" } }] }],
  });
}
firstSlide.images.add({ name: "Brand mark", position: { left: 1000, top: 36, width: 96, height: 96 }, dataUrl: ONE_PIXEL_PNG });

const authoredProfile = authored.designProfile();
assert.equal(authoredProfile.schema, "office-kit/pptx-design-profile/v1");
assert.deepEqual(authoredProfile.source, { sourceBound: false });
assert.equal(authoredProfile.canvas.aspectRatio, 1.77778);
assert.equal(authoredProfile.designLanguage.density.slides, 1);
assert.equal(authoredProfile.designLanguage.typography.fontSizePt.samples >= 2, true);
assert.equal(authoredProfile.componentCandidates.available, false);
assert.equal(authoredProfile.reusableComponents.length >= 1, true);
assert.equal(JSON.stringify(authoredProfile).includes("<p:"), false);

const fixturePath = path.resolve(import.meta.dirname, "../evals/assets/presentations/strategy-review.pptx");
const fixtureBytes = await readFile(fixturePath);
const imported = await PresentationFile.importPptx(new FileBlob(fixtureBytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }));
const importedProfile = imported.designProfile({ maxItems: 32 });
assert.equal(importedProfile.schema, "office-kit/pptx-design-profile/v1");
assert.equal(importedProfile.source.sourceBound, true);
assert.match(importedProfile.source.revisionSha256, /^[0-9a-f]{64}$/u);
assert.equal(importedProfile.canvas.width, 1280);
assert.equal(importedProfile.designLanguage.density.slides, imported.slides.count);
assert.equal(importedProfile.nativeOpaque.kinds.diagram >= 1, true);
assert.equal(importedProfile.componentCandidates.available, true);
assert.equal(importedProfile.componentCandidates.total > 0, true);
assert.equal(importedProfile.componentCandidates.inspectOnlyCandidateIds.length > 0, true);
assert.equal(importedProfile.componentCandidates.inspectOnlyCandidateIds.every((id) => /^pc_[0-9a-f]{32}$/u.test(id)), true);
assert.equal(JSON.stringify(importedProfile).includes(fixturePath), false);
const noOp = await PresentationFile.exportPptx(imported);
assert.deepEqual([...noOp.bytes], [...fixtureBytes], "design profile inspection must not mutate imported source bytes");

const withoutCandidates = imported.designProfile({ includeComponentCandidates: false });
assert.deepEqual(withoutCandidates.componentCandidates, { available: false, reason: "component candidate inspection disabled" });

console.log("presentation design profile runtime smoke ok");
