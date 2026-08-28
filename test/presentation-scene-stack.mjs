import assert from "node:assert/strict";

import { FileBlob, Presentation, PresentationFile } from "../src/index.mjs";

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
assert.throws(() => reopened.elements.getItem("photo").bringToFront(), /cannot be safely changed/);

console.log("presentation scene stack smoke ok");
