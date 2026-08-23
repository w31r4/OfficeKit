import assert from "node:assert/strict";

import JSZip from "jszip";

import {
  compareContinuationRenderedPages,
  evaluateContinuationPackageOracle,
  sha256,
} from "../scripts/pptx-programmable-import-oracle.mjs";

const contentTypes = '<Types><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>';
const sourceBytes = await makeZip({
  "[Content_Types].xml": contentTypes,
  "ppt/presentation.xml": '<p:presentation xmlns:p="urn:p" xmlns:r="urn:r"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>',
  "ppt/_rels/presentation.xml.rels": '<Relationships><Relationship Id="rId1" Type="urn/officeDocument/relationships/slide" Target="slides/slide1.xml"/></Relationships>',
  "ppt/slides/slide1.xml": '<p:sld xmlns:p="urn:p" xmlns:r="urn:r" xmlns:a="urn:a"><a:blip r:embed="rId2"/></p:sld>',
  "ppt/slides/_rels/slide1.xml.rels": '<Relationships><Relationship Id="rId2" Type="urn/officeDocument/relationships/image" Target="../media/image1.svg"/></Relationships>',
  "ppt/media/image1.svg": '<svg xmlns="http://www.w3.org/2000/svg"><text>Kimsoong</text><text>Customer Loyalty Programme</text></svg>',
  "docProps/canary.bin": Buffer.from([1, 2, 3]),
});
const source = { id: "synthetic-continuation", sha256: sha256(sourceBytes), slideCount: 1 };
const task = {
  id: "synthetic-svg-continuation",
  sourceSlide: 1,
  targetPageAfterAppend: 2,
  edits: [
    { phase: 1, operation: "svg-text", nodeId: "svg-text-1", expected: "Kimsoong", value: "OfficeKit Acceptance" },
    { phase: 2, operation: "svg-text", nodeId: "svg-text-2", expected: "Customer Loyalty Programme", value: "Programmable import verified" },
  ],
};
const outputBytes = await makeZip({
  "[Content_Types].xml": contentTypes.replace("</Types>", '<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'),
  "ppt/presentation.xml": '<p:presentation xmlns:r="urn:r" xmlns:p="urn:p"><p:sldIdLst><p:sldId r:id="rId1" id="256"/><p:sldId id="257" r:id="RnewSlide"/></p:sldIdLst></p:presentation>',
  "ppt/_rels/presentation.xml.rels": '<Relationships><Relationship Target="slides/slide1.xml" Type="urn/officeDocument/relationships/slide" Id="rId1"/><Relationship Id="RnewSlide" Type="urn/officeDocument/relationships/slide" Target="slides/slide2.xml"/></Relationships>',
  "ppt/slides/slide1.xml": '<p:sld xmlns:p="urn:p" xmlns:r="urn:r" xmlns:a="urn:a"><a:blip r:embed="rId2"/></p:sld>',
  "ppt/slides/_rels/slide1.xml.rels": '<Relationships><Relationship Id="rId2" Type="urn/officeDocument/relationships/image" Target="../media/image1.svg"/></Relationships>',
  "ppt/slides/slide2.xml": '<p:sld xmlns:a="urn:a" xmlns:r="urn:r" xmlns:p="urn:p"><a:blip r:embed="Redited" /></p:sld>',
  "ppt/slides/_rels/slide2.xml.rels": '<Relationships><Relationship Id="Redited" Type="urn/officeDocument/relationships/image" Target="../media/image2.svg"/></Relationships>',
  "ppt/media/image1.svg": '<svg xmlns="http://www.w3.org/2000/svg"><text>Kimsoong</text><text>Customer Loyalty Programme</text></svg>',
  "ppt/media/image2.svg": '<svg xmlns="http://www.w3.org/2000/svg"><text>OfficeKit Acceptance</text><text>Programmable import verified</text></svg>',
  "docProps/canary.bin": Buffer.from([1, 2, 3]),
});

const result = await evaluateContinuationPackageOracle({ sourceBytes, outputBytes, source, task });
assert.equal(result.nonTargetExistingPartsByteIdentical, true);
assert.equal(result.relationships.sourceRelationshipPartsByteIdentical, true);
assert.equal(result.targetMask.svgMask.passed, true);
assert.equal(result.addedGraph.passed, true);
assert.equal(result.outputSlideCount, 2);

const driftedOutput = await makeZip({
  ...(await unzip(outputBytes)),
  "docProps/canary.bin": Buffer.from([1, 2, 4]),
});
await assert.rejects(
  evaluateContinuationPackageOracle({ sourceBytes, outputBytes: driftedOutput, source, task }),
  /non-target existing OPC drift/u,
);

const sourceRender = { pages: [{ page: 1, sha256: "source-page" }] };
const outputRender = { pages: [{ page: 1, sha256: "source-page" }, { page: 2, sha256: "appended-page" }], cacheHit: false };
assert.equal(compareContinuationRenderedPages(sourceRender, outputRender, 2).nonTargetPagesPixelIdentical, true);
assert.throws(
  () => compareContinuationRenderedPages(sourceRender, { pages: [{ page: 1, sha256: "drift" }, outputRender.pages[1]] }, 2),
  /Non-target rendered pages changed/u,
);

async function makeZip(parts) {
  const zip = new JSZip();
  for (const [name, value] of Object.entries(parts)) zip.file(name, value, { date: new Date(0) });
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 }, platform: "UNIX" });
}

async function unzip(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const parts = {};
  for (const [name, entry] of Object.entries(zip.files)) {
    if (!entry.dir) parts[name] = await entry.async("nodebuffer");
  }
  return parts;
}

console.log("PPTX programmable-import continuation oracle ok");
