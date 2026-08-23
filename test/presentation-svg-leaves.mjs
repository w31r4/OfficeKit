import assert from "node:assert/strict";

import { Presentation, PresentationFile } from "../src/index.mjs";

function svgDataUrl(source) {
  return `data:image/svg+xml;base64,${Buffer.from(source, "utf8").toString("base64")}`;
}

function svgSource(dataUrl) {
  return Buffer.from(String(dataUrl).split(",", 2)[1], "base64").toString("utf8");
}

function leaf(image, predicate) {
  const value = image.getSvgEditLeaves().find(predicate);
  assert.ok(value, "expected SVG leaf was not issued");
  return value;
}

function edit(image, predicate, value) {
  const current = leaf(image, predicate);
  return image.editSvgLeaf(current.id, { expectedHash: current.expectedHash, value });
}

const SAFE_SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 360"><defs><linearGradient id="g"><stop offset="0" stop-color="#FFFFFF"/></linearGradient></defs><rect id="panel" fill="#abc" stroke="#123456" opacity="0.75" transform="translate(12,-4)"/><g transform="rotate(-90 220 280)"><path fill="url(#g)" transform="scale(1.5)"/></g><circle style="fill:#FFEEDD" cx="20" cy="20" r="10"/><use href="#panel"/></svg>';
const deck = Presentation.create({ slideSize: { width: 640, height: 360 } });
const slide = deck.slides.add({ name: "Safe SVG leaves" });
const image = slide.images.add({
  dataUrl: svgDataUrl(SAFE_SVG),
  position: { left: 0, top: 0, width: 640, height: 360 },
});

const capability = image.svgEditCapability;
assert.equal(capability.supported, true);
assert.match(capability.sourceSha256, /^[0-9a-f]{64}$/u);
assert.equal(capability.sourceRevisionSha256, undefined);
assert.deepEqual(
  Object.fromEntries([...new Set(capability.leaves.map((item) => item.leafKind))].sort().map((kind) => [kind, capability.leaves.filter((item) => item.leafKind === kind).length])),
  { svgFillRgb: 1, svgOpacity: 1, svgStrokeRgb: 1, svgTransformScalar: 6 },
);
assert.equal(capability.leaves.some((item) => item.attribute === "style"), false);
assert.equal(capability.leaves.some((item) => item.value === "url(#g)"), false);
assert.equal(Object.isFrozen(capability.leaves), true);
const defensiveLeaves = image.getSvgEditLeaves();
defensiveLeaves[0].value = "mutated copy";
assert.notEqual(image.getSvgEditLeaves()[0].value, "mutated copy");

// Leaf IDs are bound to the owning presentation object as well as the exact
// SVG bytes. Identical bytes in a second image cannot reuse the first image's
// issued capability.
const twin = slide.images.add({
  dataUrl: svgDataUrl(SAFE_SVG),
  position: { left: 0, top: 0, width: 64, height: 36 },
});
const firstFill = leaf(image, (item) => item.leafKind === "svgFillRgb");
assert.throws(
  () => twin.editSvgLeaf(firstFill.id, { expectedHash: firstFill.expectedHash, value: "#001122" }),
  (error) => error.code === "presentation_svg_leaf_not_issued",
);

const fillEdit = image.editSvgLeaf(firstFill.id, { expectedHash: firstFill.expectedHash, value: "#def" });
assert.equal(fillEdit.leafKind, "svgFillRgb");
assert.equal(fillEdit.oldValue, "#AABBCC");
assert.equal(fillEdit.value, "#DDEEFF");
assert.throws(
  () => image.editSvgLeaf(firstFill.id, { expectedHash: firstFill.expectedHash, value: "#001122" }),
  (error) => error.code === "presentation_svg_leaf_not_issued",
);

const currentFill = leaf(image, (item) => item.leafKind === "svgFillRgb");
assert.throws(
  () => image.editSvgLeaf(currentFill.id, { expectedHash: "0".repeat(64), value: "#001122" }),
  (error) => error.code === "presentation_svg_leaf_stale",
);
assert.throws(
  () => image.editSvgLeaf(currentFill.id, { expectedHash: currentFill.expectedHash, value: "#ddeeff" }),
  (error) => error.code === "presentation_svg_leaf_noop",
);
assert.throws(
  () => image.editSvgLeaf(currentFill.id, { expectedHash: currentFill.expectedHash, value: "red" }),
  (error) => error.code === "invalid_presentation_svg_leaf",
);

edit(image, (item) => item.leafKind === "svgStrokeRgb", "#654321");
edit(image, (item) => item.leafKind === "svgOpacity", 0.5);
edit(image, (item) => item.leafKind === "svgTransformScalar" && item.component === "x", 25);
edit(image, (item) => item.leafKind === "svgTransformScalar" && item.component === "angle", -45);
const scale = leaf(image, (item) => item.leafKind === "svgTransformScalar" && item.component === "scale");
assert.throws(
  () => image.editSvgLeaf(scale.id, { expectedHash: scale.expectedHash, value: 0 }),
  (error) => error.code === "invalid_presentation_svg_leaf",
);
assert.equal(
  svgSource(image.dataUrl),
  SAFE_SVG
    .replace('fill="#abc"', 'fill="#DDEEFF"')
    .replace('stroke="#123456"', 'stroke="#654321"')
    .replace('opacity="0.75"', 'opacity="0.5"')
    .replace('translate(12,-4)', 'translate(25,-4)')
    .replace('rotate(-90 220 280)', 'rotate(-45 220 280)'),
  "safe edits must splice only the issued scalar token",
);

for (const source of [
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><script>alert(1)</script><rect fill="#fff"/></svg>',
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><rect ONLOAD="alert(1)" fill="#fff"/></svg>',
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><style>.x{fill:#fff}</style><rect fill="#fff"/></svg>',
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><rect class="x" fill="#fff"/></svg>',
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><rect style="fill:url(https://example.invalid/a.svg)"/><path fill="#fff"/></svg>',
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><image href="https://example.invalid/a.png"/><rect fill="#fff"/></svg>',
]) {
  const unsafeDeck = Presentation.create({ slideSize: { width: 100, height: 100 } });
  const unsafe = unsafeDeck.slides.add({ name: "Unsafe SVG" }).images.add({
    dataUrl: svgDataUrl(source),
    position: { left: 0, top: 0, width: 100, height: 100 },
  });
  assert.equal(unsafe.svgEditCapability.supported, false);
  assert.throws(
    () => unsafe.editSvgLeaf("sl_invalid", { expectedHash: "0".repeat(64), value: "#FFFFFF" }),
    (error) => error.code === "unsupported_presentation_svg_leaf",
  );
}

// The capability survives PPTX export/import, gains a package revision binding,
// and remains editable without converting the SVG into ordinary PPT shapes.
const exported = await PresentationFile.exportPptx(deck);
const imported = await PresentationFile.importPptx(exported);
const importedImage = imported.slides.items[0].images.items[0];
assert.equal(importedImage.svgEditCapability.supported, true);
assert.match(importedImage.svgEditCapability.sourceRevisionSha256, /^[0-9a-f]{64}$/u);
assert.equal(svgSource(importedImage.dataUrl), svgSource(image.dataUrl));
edit(importedImage, (item) => item.leafKind === "svgStrokeRgb", "#001122");
const roundTrip = await PresentationFile.importPptx(await PresentationFile.exportPptx(imported));
const roundTripImage = roundTrip.slides.items[0].images.items[0];
assert.equal(leaf(roundTripImage, (item) => item.leafKind === "svgStrokeRgb").value, "#001122");
assert.equal(svgSource(roundTripImage.dataUrl), svgSource(importedImage.dataUrl));

console.log("presentation svg leaves smoke ok");
