import assert from "node:assert/strict";

import { FileBlob } from "../src/index.mjs";
import { Presentation, PresentationFile } from "../src/presentation/index.mjs";

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";
const pngBytes = Buffer.from(PNG.split(",")[1], "base64");

const deck = Presentation.create({ slideSize: { width: 640, height: 360 } });
const slide = deck.slides.add({ name: "mixed layers" });
const title = slide.shapes.add({
  name: "title",
  geometry: "textbox",
  position: { left: 40, top: 40, width: 520, height: 72 },
  text: "Editable foreground",
});
const scrim = slide.shapes.add({
  name: "scrim",
  geometry: "rect",
  position: { left: 0, top: 0, width: 640, height: 360 },
  fill: { color: "#000000", opacity: 0.48 },
});
const background = slide.setBackgroundImage({
  name: "photo",
  blob: new FileBlob(pngBytes, { type: "image/png" }),
  fit: "cover",
  alt: "Decorative background",
});

scrim.moveAfter(background);
title.bringToFront();
assert.deepEqual(slide.elements.items.map((element) => element.name), ["photo", "scrim", "title"]);
assert.deepEqual(slide.inspectRecords(new Set(["layer"])).map(({ stackIndex }) => stackIndex), [0, 1, 2]);

assert.throws(() => title.moveBefore({}), /ordered slide or group scene stack/);
scrim.delete();
assert.deepEqual(slide.elements.items.map((element) => element.name), ["photo", "title"]);
slide.shapes.add({ name: "scrim", geometry: "rect", position: { left: 0, top: 0, width: 640, height: 360 }, fill: { color: "#000000", opacity: 0.48 } }).moveAfter(background);
title.bringToFront();

const output = await PresentationFile.exportPptx(deck);
const reimported = await PresentationFile.importPptx(output);
const reopened = reimported.slides.items[0];
assert.deepEqual(reopened.elements.items.map((element) => element.name), ["photo", "scrim", "title"]);
assert.deepEqual(reopened.elements.getItem("scrim").fill, { color: "#000000", opacity: 0.48 });
assert.equal(reopened.elements.getItem("photo").zOrderCapability.editable, true);
reopened.elements.getItem("photo").bringToFront();
const reorderedOutput = await PresentationFile.exportPptx(reimported);
const reordered = await PresentationFile.importPptx(reorderedOutput);
assert.deepEqual(reordered.slides.items[0].elements.items.map((element) => element.name), ["scrim", "title", "photo"]);

// A real NASA fixture exposed that source-bound reorder could lose imported
// picture-bullet assets even though the source package itself was preserved.
// Keep one compact import -> edit -> export regression beside the scene-stack
// contract instead of adding a fixture matrix.
const bulletDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
const bulletSlide = bulletDeck.slides.add({ name: "picture bullet" });
bulletSlide.shapes.add({
  name: "list",
  geometry: "textbox",
  position: { left: 40, top: 40, width: 520, height: 72 },
  text: [{ runs: [{ text: "Preserved marker" }], bulletImage: { dataUrl: PNG } }],
});
bulletSlide.shapes.add({ name: "accent", geometry: "rect", position: { left: 32, top: 32, width: 8, height: 88 }, fill: "#0f766e" });
const importedBulletDeck = await PresentationFile.importPptx(await PresentationFile.exportPptx(bulletDeck));
importedBulletDeck.slides.items[0].elements.getItem("accent").sendToBack();
const reopenedBulletDeck = await PresentationFile.importPptx(await PresentationFile.exportPptx(importedBulletDeck));
assert.equal(reopenedBulletDeck.slides.items[0].elements.getItem("list").text.paragraphs[0].bulletImage.dataUrl, PNG);

console.log("presentation scene stack smoke ok");
