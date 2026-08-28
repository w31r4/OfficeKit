import assert from "node:assert/strict";
import JSZip from "jszip";

import { Presentation, PresentationFile } from "../src/index.mjs";

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

console.log("presentation CJK font smoke ok");
