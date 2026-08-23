import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

const evidence = JSON.parse(await readFile(path.resolve("evals/pptx-lossless/source-continuation-native.v2.json"), "utf8"));
assert.equal(evidence.schema, "office-kit/pptx-source-continuation-native-evidence/v2");
assert.deepEqual(evidence.renderer, { office: "LibreOffice", raster: "Poppler", dpi: 120 });
assert.equal(evidence.sources.length, 3);
for (const source of evidence.sources) {
  assert.equal(source.outputSlideCount, source.sourceSlideCount + 1);
  assert.equal(source.insertedSlide, source.sourceSlideCount + 1);
  assert.equal(source.continuationKind, "bounded-overlay");
  assert.equal(source.nonTargetPagesPixelIdentical, true);
  assert.equal(source.insertedPageRendered, true);
  assert.equal(source.insertedPageChangedFromClone, true);
  assert.match(source.insertedPageChange.cloneHash, /^[a-f0-9]{64}$/u);
  assert.match(source.insertedPageChange.outputHash, /^[a-f0-9]{64}$/u);
  assert.notEqual(source.insertedPageChange.outputHash, source.insertedPageChange.cloneHash);
  assert.ok(source.insertedPageChange.differentPixels > 0);
  assert.ok(source.insertedPageChange.mismatchRatio > 0 && source.insertedPageChange.mismatchRatio < 1);
  assert.equal(source.target.kind, "bounded-overlay");
  assert.equal(source.target.text.name, "officekit-source-derived-text");
  assert.equal(source.target.accent.name, "officekit-source-derived-accent");
  assert.equal(source.target.image.name, "officekit-source-derived-image");
  assert.equal(source.pages.length, source.sourceSlideCount);
  for (const page of source.pages) {
    assert.equal(page.pixelIdentical, true);
    assert.match(page.sourceHash, /^[a-f0-9]{64}$/u);
    assert.match(page.outputHash, /^[a-f0-9]{64}$/u);
    assert.equal(page.outputSlide, page.sourceSlide);
  }
}
console.log("pptx source continuation native render ok");
