import assert from "node:assert/strict";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import JSZip from "jszip";
import sharp from "sharp";

import {
  canonicalizeXml,
  compareRenderedPages,
  evaluatePackageOracle,
  sha256,
} from "../scripts/pptx-programmable-import-oracle.mjs";

assert.equal(
  canonicalizeXml('<?xml version="1.0"?><a:x xmlns:a="urn:a" x="1"><a:y /></a:x>'),
  canonicalizeXml('<?xml version="1.0" encoding="utf-8"?>\n<a:x x="1" xmlns:a="urn:a">\n<a:y/>\n</a:x>'),
);

const source = {
  id: "synthetic",
  sha256: null,
};
const nativeIntent = {
  id: "title",
  operation: "native-leaf",
  leafKind: "text",
  expected: "Before",
  value: "After",
  oracle: { changedParts: ["ppt/slides/slide1.xml"] },
};
const sourceBytes = await makeZip({
  "[Content_Types].xml": "<Types/>",
  "ppt/slides/slide1.xml": "<p:sld xmlns:p=\"urn:p\"><a:t xmlns:a=\"urn:a\">Before</a:t></p:sld>",
  "ppt/slides/_rels/slide1.xml.rels": "<Relationships/>",
  "ppt/media/canary.bin": Buffer.from([1, 2, 3]),
});
source.sha256 = sha256(sourceBytes);
const validOutput = await makeZip({
  "[Content_Types].xml": "<Types/>",
  "ppt/slides/slide1.xml": "<p:sld xmlns:p=\"urn:p\"><a:t xmlns:a=\"urn:a\">After</a:t></p:sld>",
  "ppt/slides/_rels/slide1.xml.rels": "<Relationships/>",
  "ppt/media/canary.bin": Buffer.from([1, 2, 3]),
});
const valid = await evaluatePackageOracle({ sourceBytes, outputBytes: validOutput, source, intent: nativeIntent });
assert.equal(valid.nonTargetPartsByteIdentical, true);
assert.equal(valid.relationships.passed, true);
assert.equal(valid.targetMask.passed, true);

const driftOutput = await makeZip({
  "[Content_Types].xml": "<Types/>",
  "ppt/slides/slide1.xml": "<p:sld xmlns:p=\"urn:p\"><a:t xmlns:a=\"urn:a\">After</a:t></p:sld>",
  "ppt/slides/_rels/slide1.xml.rels": "<Relationships/>",
  "ppt/media/canary.bin": Buffer.from([1, 2, 4]),
});
await assert.rejects(
  evaluatePackageOracle({ sourceBytes, outputBytes: driftOutput, source, intent: nativeIntent }),
  /changed OPC parts/u,
);

const relSource = await makeZip({
  "[Content_Types].xml": "<Types/>",
  "ppt/slides/slide1.xml": '<p:sld xmlns:p="urn:p" xmlns:r="urn:r"><a:blip xmlns:a="urn:a" r:embed="rId2"/></p:sld>',
  "ppt/slides/_rels/slide1.xml.rels": '<Relationships><Relationship Id="rId1" Type="urn/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="urn/image" Target="../media/image1.svg"/></Relationships>',
  "ppt/media/image1.svg": '<svg xmlns="http://www.w3.org/2000/svg"><text>Old SVG</text></svg>',
});
const relSourceRecord = { id: "svg", sha256: sha256(relSource) };
const relIntent = {
  id: "svg-edit",
  operation: "svg-text",
  expected: "Old SVG",
  value: "New SVG",
  oracle: {
    sourceSvgPart: "ppt/media/image1.svg",
    changedParts: ["ppt/slides/slide1.xml", "ppt/slides/_rels/slide1.xml.rels"],
    addedParts: ["ppt/media/image2.svg"],
  },
};
const relOutput = await makeZip({
  "[Content_Types].xml": "<Types/>",
  "ppt/slides/slide1.xml": '<p:sld xmlns:r="urn:r" xmlns:p="urn:p"><a:blip r:embed="Rnew" xmlns:a="urn:a" /></p:sld>',
  "ppt/slides/_rels/slide1.xml.rels": '<Relationships><Relationship Target="../slideLayouts/slideLayout1.xml" Type="urn/slideLayout" Id="rId1"/><Relationship Target="../media/image1.svg" Type="urn/image" Id="rId2"/><Relationship Id="Rnew" Type="urn/image" Target="/ppt/media/image2.svg"/></Relationships>',
  "ppt/media/image1.svg": '<svg xmlns="http://www.w3.org/2000/svg"><text>Old SVG</text></svg>',
  "ppt/media/image2.svg": '<svg xmlns="http://www.w3.org/2000/svg"><text>New SVG</text></svg>',
});
const svgResult = await evaluatePackageOracle({ sourceBytes: relSource, outputBytes: relOutput, source: relSourceRecord, intent: relIntent });
assert.equal(svgResult.targetMask.passed, true);
assert.equal(svgResult.relationships.existingRelationshipsPreserved, true);

const temporary = await mkdtemp(path.join(os.tmpdir(), "officekit-pptx-pixel-oracle-"));
const sourcePages = [];
const outputPages = [];
for (let page = 1; page <= 3; page += 1) {
  const before = await sharp({ create: { width: 8, height: 8, channels: 4, background: page === 2 ? "#ff0000" : "#ffffff" } }).png().toBuffer();
  const after = await sharp({ create: { width: 8, height: 8, channels: 4, background: page === 2 ? "#0000ff" : "#ffffff" } }).png().toBuffer();
  await writeFile(path.join(temporary, `source-${page}.png`), before);
  await writeFile(path.join(temporary, `output-${page}.png`), after);
  sourcePages.push({ page, sha256: sha256(await readFile(path.join(temporary, `source-${page}.png`))) });
  outputPages.push({ page, sha256: sha256(await readFile(path.join(temporary, `output-${page}.png`))) });
}
const pixel = compareRenderedPages({ pages: sourcePages }, { pages: outputPages, cacheHit: false }, 2);
assert.equal(pixel.nonTargetPagesPixelIdentical, true);
assert.equal(pixel.targetPageChanged, true);
assert.throws(() => compareRenderedPages({ pages: sourcePages }, { pages: outputPages, cacheHit: false }, 1), /Non-target rendered pages changed/u);

async function makeZip(parts) {
  const zip = new JSZip();
  for (const [name, value] of Object.entries(parts)) zip.file(name, value, { date: new Date(0) });
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 }, platform: "UNIX" });
}

console.log("PPTX programmable-import independent oracle ok");
