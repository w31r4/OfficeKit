import assert from "node:assert/strict";
import JSZip from "jszip";

import { Presentation, PresentationFile } from "../src/index.mjs";

const layoutDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
const layoutSlide = layoutDeck.slides.add({ name: "Decorative overlap" });
const orbit = layoutSlide.shapes.add({
  name: "orbit",
  geometry: "ellipse",
  position: { left: -48, top: 128, width: 220, height: 220 },
  fill: "transparent",
  line: { fill: "#E2E8F0", width: 8 },
  accessibility: { decorative: true },
});
layoutSlide.shapes.add({
  name: "evidence",
  geometry: "textbox",
  position: { left: 32, top: 160, width: 280, height: 48 },
  text: "Evidence",
});
const layout = layoutSlide.validateLayout();
assert.equal(layout.ok, true);
assert.equal(layout.issues.some((issue) => issue.type === "overlap" && issue.ids?.includes(orbit.id)), false);
assert.equal(layout.issues.some((issue) => issue.id === orbit.id), false);

const cjkDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
cjkDeck.slides.add().shapes.add({
  geometry: "textbox",
  text: "中文字体不会丢失",
  position: { left: 40, top: 80, width: 520, height: 80 },
  textStyle: { fontFamily: "PingFang SC", fontSize: 36 },
});
const pptx = await PresentationFile.exportPptx(cjkDeck);
const zip = await JSZip.loadAsync(pptx.bytes);
const slideXml = await zip.file("ppt/slides/slide1.xml").async("text");
assert.match(slideXml, /<a:ea\b[^>]*typeface="PingFang SC"/u);
const roundTrip = await PresentationFile.importPptx(pptx);
assert.equal(roundTrip.slides.getItem(0).shapes.items[0].text.paragraphs[0].runs[0].style.fontFamilyEastAsia, "PingFang SC");

console.log("presentation CJK font and decorative layout smoke ok");
