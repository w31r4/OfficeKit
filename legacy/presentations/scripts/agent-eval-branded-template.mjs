import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import zlib from "node:zlib";

import JSZip from "jszip";

import {
  FileBlob,
  Presentation,
  PresentationFile,
  SpreadsheetFile,
  Workbook,
} from "../src/index.mjs";

export const BRANDED_TEMPLATE_FIXTURE = Object.freeze({
  presentationName: "quarterly-board-template.pptx",
  replacementImageName: "replacement-product.png",
  slideCount: 8,
  titleSlideIndex: 1,
  titleShapeName: "board-title-target",
  originalTitle: "Quarterly Board Review — Pending",
  replacementTitle: "Quarterly Board Review — Approved",
  tableSlideIndex: 4,
  tableName: "board-metrics-table",
  tableOriginalValues: Object.freeze(["120", "130", "140"]),
  tableReplacementValues: Object.freeze(["125", "135", "145"]),
  chartSlideIndex: 6,
  chartName: "board-revenue-chart",
  chartOriginalTitle: "Revenue by region",
  chartReplacementTitle: "Revenue outlook",
  chartOriginalValues: Object.freeze([42, 56, 63]),
  chartReplacementValues: Object.freeze([45, 60, 70]),
  imageSlideIndex: 7,
  imageName: "product-image-target",
  imageAlt: "Quarterly product launch image",
  imagePosition: Object.freeze({ left: 116, top: 178, width: 420, height: 292 }),
  smartArtSlideIndex: 3,
  smartArtPartPaths: Object.freeze([
    "ppt/diagrams/brand-data.xml",
    "ppt/diagrams/brand-layout.xml",
    "ppt/diagrams/brand-style.xml",
    "ppt/diagrams/brand-colors.xml",
  ]),
  oleSlideIndex: 5,
  oleWorkbookPath: "ppt/embeddings/brand-finance.xlsx",
  olePreviewPath: "ppt/media/brand-finance-preview.png",
  oleWorkbookRelationshipId: "rIdBrandFinanceWorkbook",
  olePreviewRelationshipId: "rIdBrandFinancePreview",
  customShowName: "Board route",
  customShowNativeId: 41,
  notesSlideIndex: 2,
  notesText: "Review the approved controls and retain the decision evidence.",
  commentText: "Confirm the final board evidence before circulation.",
});

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function pngChunk(type, payload) {
  const typeBytes = Buffer.from(type, "ascii");
  const body = Buffer.concat([typeBytes, payload]);
  const checksum = Buffer.alloc(4);
  checksum.writeUInt32BE(crc32(body), 0);
  const length = Buffer.alloc(4);
  length.writeUInt32BE(payload.length, 0);
  return Buffer.concat([length, body, checksum]);
}

function solidPng(width, height, rgb) {
  const rows = [];
  for (let y = 0; y < height; y += 1) {
    const row = Buffer.alloc(1 + width * 4);
    row[0] = 0;
    for (let x = 0; x < width; x += 1) {
      const offset = 1 + x * 4;
      row[offset] = rgb[0];
      row[offset + 1] = rgb[1];
      row[offset + 2] = rgb[2];
      row[offset + 3] = 255;
    }
    rows.push(row);
  }
  const header = Buffer.alloc(13);
  header.writeUInt32BE(width, 0);
  header.writeUInt32BE(height, 4);
  header[8] = 8;
  header[9] = 6;
  header[10] = 0;
  header[11] = 0;
  header[12] = 0;
  return Buffer.concat([
    Buffer.from("89504e470d0a1a0a", "hex"),
    pngChunk("IHDR", header),
    pngChunk("IDAT", zlib.deflateSync(Buffer.concat(rows), { level: 9 })),
    pngChunk("IEND", Buffer.alloc(0)),
  ]);
}

export function brandedReplacementPng() {
  return solidPng(160, 96, [14, 116, 144]);
}

function dataUrl(bytes, mime = "image/png") {
  return `data:${mime};base64,${Buffer.from(bytes).toString("base64")}`;
}

function xmlEscape(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function smartArtParts() {
  return [
    {
      path: BRANDED_TEMPLATE_FIXTURE.smartArtPartPaths[0],
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
      xml: '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="{A1111111-1111-4111-8111-111111111111}" type="doc"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Align</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="{B2222222-2222-4222-8222-222222222222}" type="doc"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Execute</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="{C3333333-3333-4333-8333-333333333333}" type="doc"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Measure</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst/><dgm:bg/><dgm:whole/></dgm:dataModel>',
    },
    {
      path: BRANDED_TEMPLATE_FIXTURE.smartArtPartPaths[1],
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
      xml: '<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:office-kit:brand-layout"><dgm:title val="Brand operating model"/><dgm:desc val="Self-authored locked layout"/><dgm:catLst/><dgm:layoutNode name="root"/></dgm:layoutDef>',
    },
    {
      path: BRANDED_TEMPLATE_FIXTURE.smartArtPartPaths[2],
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
      xml: '<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:office-kit:brand-style"><dgm:title val="Brand style"/><dgm:desc val="Self-authored locked style"/><dgm:catLst/><dgm:styleLbl name="brand"/></dgm:styleDef>',
    },
    {
      path: BRANDED_TEMPLATE_FIXTURE.smartArtPartPaths[3],
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
      xml: '<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:office-kit:brand-colors"><dgm:title val="Brand colors"/><dgm:desc val="Self-authored locked colors"/><dgm:catLst/></dgm:colorsDef>',
    },
  ];
}

async function embeddedWorkbook() {
  const workbook = Workbook.create({ name: "Brand finance" });
  const sheet = workbook.worksheets.add("Finance");
  sheet.getRange("A1:B4").values = [["Metric", "Value"], ["Revenue", 145], ["Margin", 0.46], ["Status", "Approved"]];
  return SpreadsheetFile.exportXlsx(workbook);
}

function oleFrame() {
  const fixture = BRANDED_TEMPLATE_FIXTURE;
  return `<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="180" name="Embedded finance workbook"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="1800000"/><a:ext cx="3657600" cy="1828800"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj showAsIcon="1" r:id="${fixture.oleWorkbookRelationshipId}" imgW="965200" imgH="609600" progId="Excel.Sheet.12"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${fixture.olePreviewRelationshipId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="1800000"/><a:ext cx="3657600" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>`;
}

function smartArtFrame() {
  return '<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="220" name="Brand operating SmartArt"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="1800000"/><a:ext cx="5486400" cy="2743200"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/officeDocument/2006/diagram" r:dm="rIdBrandDiagramData" r:lo="rIdBrandDiagramLayout" r:qs="rIdBrandDiagramStyle" r:cs="rIdBrandDiagramColors"/></a:graphicData></a:graphic></p:graphicFrame>';
}

function relationshipXml(entries) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries.map((entry) => `<Relationship Id="${entry.id}" Type="${entry.type}" Target="${entry.target}"/>`).join("")}</Relationships>`;
}

export async function generateBrandedTemplate(target, imageTarget) {
  const fixture = BRANDED_TEMPLATE_FIXTURE;
  const sourceImage = solidPng(160, 96, [30, 64, 112]);
  const replacementImage = brandedReplacementPng();
  const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 }, commentFormat: "modern" });

  const addTitle = (slide, name, text, color = "#0F172A") => {
    const title = slide.shapes.add({ name, geometry: "textbox", position: { left: 72, top: 58, width: 1080, height: 82 }, text, fill: "none", line: { style: "solid", fill: "none", width: 0 } });
    title.text.style = { fontSize: 32, bold: true, color };
    return title;
  };
  const addCanary = (slide, text = "All non-target brand and package objects remain unchanged.") => {
    const canary = slide.shapes.add({ name: "preservation-canary", geometry: "textbox", position: { left: 72, top: 620, width: 1100, height: 42 }, text, fill: "none", line: { style: "solid", fill: "none", width: 0 } });
    canary.text.style = { fontSize: 14, color: "#475569" };
  };

  const cover = presentation.slides.add({ name: "Board cover" });
  cover.setBackground({ fill: "#0F2742", mode: "solid" });
  addTitle(cover, "cover-title", "Northwind | Quarterly Board Review", "#FFFFFF");
  const coverSub = cover.shapes.add({ name: "cover-subtitle", geometry: "textbox", position: { left: 72, top: 190, width: 800, height: 64 }, text: "FY27 operating decision packet", fill: "none", line: { style: "solid", fill: "none", width: 0 } });
  coverSub.text.style = { fontSize: 22, color: "#BAE6FD" };
  addCanary(cover, "Self-authored OfficeKit PromptBench source; do not overwrite.");

  const titleSlide = presentation.slides.add({ name: "Quarterly decision" });
  titleSlide.setBackground({ fill: "#E0F2FE", mode: "solid" });
  addTitle(titleSlide, fixture.titleShapeName, fixture.originalTitle);
  const titleCopy = titleSlide.shapes.add({ name: "title-supporting-copy", geometry: "textbox", position: { left: 72, top: 190, width: 900, height: 100 }, text: "The requested local title edit is the only editable content on this slide.", fill: "none", line: { style: "solid", fill: "none", width: 0 } });
  titleCopy.text.style = { fontSize: 20, color: "#334155" };
  addCanary(titleSlide);

  const notesSlide = presentation.slides.add({ name: "Review evidence" });
  notesSlide.setBackground({ fill: "#F1F5F9", mode: "solid" });
  addTitle(notesSlide, "notes-title", "Review evidence", "#1E293B");
  const commentAnchor = notesSlide.shapes.add({ id: "brand-comment-anchor", name: "Board evidence anchor", geometry: "rect", position: { left: 72, top: 200, width: 680, height: 120 }, text: "Confirm the final board evidence before circulation.", fill: "#DBEAFE", line: { fill: "#2563EB", width: 1 } });
  notesSlide.addNotes(fixture.notesText);
  notesSlide.comments.addThread({ textMatch: { element: commentAnchor, query: fixture.commentText, occurrence: 0 } }, fixture.commentText, { id: "{11111111-1111-4111-8111-111111111111}", author: "Board reviewer", created: "2026-07-22T09:00:00Z", nativeFormat: "modern", position: { x: 1200000, y: 2100000, unit: "emu" }, comments: [{ nativeId: "{22222222-2222-4222-8222-222222222222}", author: "Finance reviewer", text: "The forecast evidence is attached.", created: "2026-07-22T09:05:00Z" }] });
  notesSlide.setTransition({ effect: "fade", speed: "medium", advanceOnClick: true });
  addCanary(notesSlide);

  const smartArtSlide = presentation.slides.add({ name: "Operating model" });
  smartArtSlide.setBackground({ fill: "#ECFDF5", mode: "solid" });
  addTitle(smartArtSlide, "smartart-title", "Operating model", "#14532D");
  smartArtSlide.shapes.add({ name: "smartart-caption", geometry: "textbox", position: { left: 72, top: 190, width: 900, height: 70 }, text: "Align → Execute → Measure", fill: "none", line: { style: "solid", fill: "none", width: 0 } });
  addCanary(smartArtSlide);

  const tableSlide = presentation.slides.add({ name: "Performance table" });
  tableSlide.setBackground({ fill: "#FFF7ED", mode: "solid" });
  addTitle(tableSlide, "table-title", "Performance table", "#7C2D12");
  tableSlide.tables.add({ name: fixture.tableName, position: { left: 72, top: 180, width: 780, height: 300 }, values: [["Metric", "Q1", "Q2", "Q3"], ["Revenue", ...fixture.tableOriginalValues], ["Margin", "42%", "44%", "46%"], ["Risk", "Amber", "Amber", "Green"]], styleOptions: { headerRow: true, bandedRows: true } });
  addCanary(tableSlide);

  const oleSlide = presentation.slides.add({ name: "Finance evidence" });
  oleSlide.setBackground({ fill: "#FDF4FF", mode: "solid" });
  addTitle(oleSlide, "ole-title", "Finance evidence", "#701A75");
  oleSlide.shapes.add({ name: "ole-caption", geometry: "textbox", position: { left: 72, top: 540, width: 960, height: 58 }, text: "Embedded finance workbook is source-owned and must remain opaque.", fill: "none", line: { style: "solid", fill: "none", width: 0 } });
  addCanary(oleSlide);

  const chartSlide = presentation.slides.add({ name: "Revenue chart" });
  chartSlide.setBackground({ fill: "#EFF6FF", mode: "solid" });
  addTitle(chartSlide, "chart-slide-title", "Revenue chart", "#1E3A8A");
  chartSlide.charts.add("bar", { name: fixture.chartName, position: { left: 72, top: 168, width: 920, height: 390 }, title: fixture.chartOriginalTitle, categories: ["North", "South", "West"], series: [{ name: "Revenue", values: [...fixture.chartOriginalValues], fill: "#0E7490" }], axes: { category: { title: "Region" }, value: { title: "Revenue", min: 0, max: 80, majorUnit: 20 } }, legend: false, dataLabels: { showValue: true, position: "top" } });
  addCanary(chartSlide);

  const imageSlide = presentation.slides.add({ name: "Product launch" });
  imageSlide.setBackground({ fill: "#F8FAFC", mode: "solid" });
  addTitle(imageSlide, "image-slide-title", "Product launch", "#0F172A");
  imageSlide.images.add({ name: fixture.imageName, alt: fixture.imageAlt, position: fixture.imagePosition, fit: "stretch", dataUrl: dataUrl(sourceImage) });
  const imageCopy = imageSlide.shapes.add({ name: "image-caption", geometry: "textbox", position: { left: 600, top: 218, width: 500, height: 180 }, text: "Replace the product image only. Preserve frame, alt text, and every other slide object.", fill: "none", line: { style: "solid", fill: "none", width: 0 } });
  imageCopy.text.style = { fontSize: 20, color: "#334155" };
  addCanary(imageSlide);

  presentation.customShows.add({ name: fixture.customShowName, nativeId: fixture.customShowNativeId, slides: [titleSlide, tableSlide, chartSlide, imageSlide] });
  const verification = presentation.verify({ visualQa: true });
  if (!verification.ok) throw new Error(`Generated branded template failed model verification: ${verification.ndjson}`);
  const exported = await PresentationFile.exportPptx(presentation);
  const workbook = await embeddedWorkbook();
  const sourceZip = await JSZip.loadAsync(exported.bytes);
  const [slide4Xml, slide4Rels, slide6Xml, slide6Rels] = await Promise.all([
    sourceZip.file("ppt/slides/slide4.xml").async("text"),
    sourceZip.file("ppt/slides/_rels/slide4.xml.rels").async("text"),
    sourceZip.file("ppt/slides/slide6.xml").async("text"),
    sourceZip.file("ppt/slides/_rels/slide6.xml.rels").async("text"),
  ]);
  const smartRels = [
    { id: "rIdBrandDiagramData", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData", target: "../diagrams/brand-data.xml" },
    { id: "rIdBrandDiagramLayout", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout", target: "../diagrams/brand-layout.xml" },
    { id: "rIdBrandDiagramStyle", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle", target: "../diagrams/brand-style.xml" },
    { id: "rIdBrandDiagramColors", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors", target: "../diagrams/brand-colors.xml" },
  ];
  const oleRels = [
    { id: fixture.oleWorkbookRelationshipId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package", target: "../embeddings/brand-finance.xlsx" },
    { id: fixture.olePreviewRelationshipId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target: "../media/brand-finance-preview.png" },
  ];
  const patches = [
    { path: "ppt/slides/slide4.xml", xml: slide4Xml.replace("</p:spTree>", `${smartArtFrame()}</p:spTree>`) },
    { path: "ppt/slides/_rels/slide4.xml.rels", xml: relationshipXml([...smartRels, ...[...slide4Rels.matchAll(/<Relationship\b[^>]*\/>/g)].map((match) => ({ id: /\bId="([^"]+)"/.exec(match[0])?.[1], type: /\bType="([^"]+)"/.exec(match[0])?.[1], target: /\bTarget="([^"]+)"/.exec(match[0])?.[1] }))].filter((entry) => entry.id && entry.type && entry.target)) },
    ...smartArtParts().map((part) => ({ path: part.path, xml: part.xml, contentType: part.contentType })),
    { path: "ppt/slides/slide6.xml", xml: slide6Xml.replace("</p:spTree>", `${oleFrame()}</p:spTree>`) },
    { path: "ppt/slides/_rels/slide6.xml.rels", xml: relationshipXml([...oleRels, ...[...slide6Rels.matchAll(/<Relationship\b[^>]*\/>/g)].map((match) => ({ id: /\bId="([^"]+)"/.exec(match[0])?.[1], type: /\bType="([^"]+)"/.exec(match[0])?.[1], target: /\bTarget="([^"]+)"/.exec(match[0])?.[1] }))].filter((entry) => entry.id && entry.type && entry.target)) },
    { path: fixture.oleWorkbookPath, bytes: workbook.bytes, contentType: XLSX_MIME },
    { path: fixture.olePreviewPath, bytes: solidPng(32, 24, [125, 75, 140]), contentType: "image/png" },
  ];
  const patched = await PresentationFile.patchPptx(exported, patches, { validate: true });
  const bytes = new Uint8Array(await patched.arrayBuffer());
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, bytes);
  if (imageTarget) {
    await fs.mkdir(path.dirname(imageTarget), { recursive: true });
    await fs.writeFile(imageTarget, replacementImage);
  }
  return { path: target, imagePath: imageTarget, type: PPTX_MIME, sha256: sha256(bytes), imageSha256: sha256(replacementImage) };
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const target = process.argv[2] || path.join("evals", "assets", "presentations", BRANDED_TEMPLATE_FIXTURE.presentationName);
  const imageTarget = process.argv[3] || path.join("evals", "assets", "presentations", BRANDED_TEMPLATE_FIXTURE.replacementImageName);
  console.log(JSON.stringify(await generateBrandedTemplate(target, imageTarget), null, 2));
}
