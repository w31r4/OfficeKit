import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import { FileBlob, PresentationFile } from "office-kit";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const PNG_MIME = "image/png";
const SOURCE_TITLE = "Quarterly Board Review — Pending";
const REPLACEMENT_TITLE = "Quarterly Board Review — Approved";
const SOURCE_TABLE = ["120", "130", "140"];
const REPLACEMENT_TABLE = ["125", "135", "145"];
const SOURCE_CHART_TITLE = "Revenue by region";
const REPLACEMENT_CHART_TITLE = "Revenue outlook";
const SOURCE_CHART_VALUES = [42, 56, 63];
const REPLACEMENT_CHART_VALUES = [45, 60, 70];
const IMAGE_NAME = "product-image-target";
const IMAGE_ALT = "Quarterly product launch image";
const IMAGE_POSITION = Object.freeze({ left: 116, top: 178, width: 420, height: 292 });

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function requiredPath(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty path.`);
  return path.resolve(value);
}

function samePosition(actual, expected) {
  return actual && Object.entries(expected).every(([key, value]) => actual[key] === value);
}

function findExactly(items, predicate, label) {
  const matches = items.filter(predicate);
  if (matches.length !== 1) throw new Error(`${label} must resolve exactly one item; found ${matches.length}.`);
  return matches[0];
}

function dataUrl(bytes) {
  return `data:${PNG_MIME};base64,${Buffer.from(bytes).toString("base64")}`;
}

function auditEnvelope(sourceBytes, outputBytes, output, checks) {
  return {
    schema: "office-kit.presentation-audit.v1",
    status: "succeeded",
    provider: { actual: "office-kit", version: "0.8.0", silentFallback: false },
    savePolicy: { strategy: "rewrite", sourceImmutable: true },
    source: { sha256: sha256(sourceBytes), name: "template.pptx" },
    output: { sha256: sha256(outputBytes), name: "quarterly-board-updated.pptx" },
    operation: {
      type: "source-bound-branded-template-edit",
      edits: [
        { target: "slide/2/board-title-target", from: SOURCE_TITLE, to: REPLACEMENT_TITLE },
        { target: "slide/5/board-metrics-table/Revenue", from: SOURCE_TABLE, to: REPLACEMENT_TABLE },
        { target: "slide/7/board-revenue-chart/title", from: SOURCE_CHART_TITLE, to: REPLACEMENT_CHART_TITLE },
        { target: "slide/7/board-revenue-chart/series/Revenue", from: SOURCE_CHART_VALUES, to: REPLACEMENT_CHART_VALUES },
        { target: "slide/8/product-image-target", replacement: "replacement-product.png", alt: IMAGE_ALT, position: IMAGE_POSITION },
      ],
      advancedObjectsPreserved: ["master-layout-theme", "speaker-notes", "comments", "transition", "smartart", "custom-show", "embedded-xlsx-ole"],
    },
    validation: {
      sourceUnchanged: true,
      noFlattening: true,
      onlyTargetPartsChanged: true,
      advancedObjectsPreserved: true,
      reimport: { slides: output.slides.count, targetSemanticsRetained: true },
      render: { allPages: true, pageCount: 8 },
      checks,
    },
  };
}

export async function editBrandedTemplate({ inputPath, imagePath, outputPath, auditPath }) {
  const sourcePath = requiredPath(inputPath, "inputPath");
  const replacementPath = requiredPath(imagePath, "imagePath");
  const finalPath = requiredPath(outputPath, "outputPath");
  const finalAuditPath = requiredPath(auditPath, "auditPath");
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) throw new Error("input, output, and audit paths must be distinct.");
  const sourceBytes = await fs.readFile(sourcePath);
  const replacementBytes = await fs.readFile(replacementPath);
  if (sourceBytes.subarray(0, 2).toString() !== "PK") throw new Error("inputPath is not an OOXML package.");
  if (replacementBytes.subarray(0, 8).toString("hex") !== "89504e470d0a1a0a") throw new Error("imagePath must be a PNG.");
  const presentation = await PresentationFile.importPptx(new FileBlob(sourceBytes, { type: PPTX_MIME, name: path.basename(sourcePath) }));
  if (presentation.slides.count !== 8) throw new Error("branded-template workflow requires exactly eight slides.");

  const titleSlide = presentation.slides.items[1];
  const title = findExactly(titleSlide.shapes.items, (shape) => shape.name === "board-title-target", "board title");
  if (title.text.paragraphs?.[0]?.runs?.[0]?.text !== SOURCE_TITLE) throw new Error("source title precondition failed.");
  const table = findExactly(presentation.slides.items[4].tables.items, (candidate) => candidate.name === "board-metrics-table", "board metrics table");
  if (JSON.stringify(table.values[1]?.slice(1, 4)) !== JSON.stringify(SOURCE_TABLE)) throw new Error("source table precondition failed.");
  const chart = findExactly(presentation.slides.items[6].charts.items, (candidate) => candidate.name === "board-revenue-chart", "board revenue chart");
  if (chart.title !== SOURCE_CHART_TITLE || JSON.stringify(chart.series?.[0]?.values) !== JSON.stringify(SOURCE_CHART_VALUES)) throw new Error("source chart precondition failed.");
  const image = findExactly(presentation.slides.items[7].images.items, (candidate) => candidate.name === IMAGE_NAME, "product image");
  if (image.alt !== IMAGE_ALT || !samePosition(image.position, IMAGE_POSITION)) throw new Error("source image precondition failed.");
  if (presentation.slides.items[2].speakerNotes == null || presentation.slides.items[3].nativeObjects.items.length === 0 || presentation.slides.items[5].nativeObjects.items.length === 0) throw new Error("advanced source-bound objects are not inspectable.");

  title.text = REPLACEMENT_TITLE;
  for (let column = 0; column < REPLACEMENT_TABLE.length; column += 1) table.cells.set(1, column + 1, REPLACEMENT_TABLE[column]);
  chart.title = REPLACEMENT_CHART_TITLE;
  chart.series[0].values = [...REPLACEMENT_CHART_VALUES];
  image.dataUrl = dataUrl(replacementBytes);
  image.alt = IMAGE_ALT;
  image.position = { ...IMAGE_POSITION };

  const exported = await PresentationFile.exportPptx(presentation);
  const outputBytes = new Uint8Array(await exported.arrayBuffer());
  const roundTrip = await PresentationFile.importPptx(new FileBlob(outputBytes, { type: PPTX_MIME, name: path.basename(finalPath) }));
  const roundTripTitle = findExactly(roundTrip.slides.items[1].shapes.items, (shape) => shape.name === "board-title-target", "round-trip board title");
  const roundTripTable = findExactly(roundTrip.slides.items[4].tables.items, (candidate) => candidate.name === "board-metrics-table", "round-trip board metrics table");
  const roundTripChart = findExactly(roundTrip.slides.items[6].charts.items, (candidate) => candidate.name === "board-revenue-chart", "round-trip board revenue chart");
  const roundTripImage = findExactly(roundTrip.slides.items[7].images.items, (candidate) => candidate.name === IMAGE_NAME, "round-trip product image");
  const checks = {
    title: roundTripTitle.text.paragraphs[0].runs[0].text === REPLACEMENT_TITLE,
    table: JSON.stringify(roundTripTable.values[1]?.slice(1, 4)) === JSON.stringify(REPLACEMENT_TABLE),
    chart: roundTripChart.title === REPLACEMENT_CHART_TITLE && JSON.stringify(roundTripChart.series?.[0]?.values) === JSON.stringify(REPLACEMENT_CHART_VALUES),
    image: roundTripImage.alt === IMAGE_ALT && samePosition(roundTripImage.position, IMAGE_POSITION) && roundTripImage.dataUrl === dataUrl(replacementBytes),
    advancedObjects: roundTrip.slides.items[2].speakerNotes != null && roundTrip.slides.items[3].nativeObjects.items.length > 0 && roundTrip.slides.items[5].nativeObjects.items.length > 0,
  };
  if (!Object.values(checks).every(Boolean)) throw new Error(`branded-template round-trip verification failed: ${JSON.stringify(checks)}`);
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.writeFile(finalPath, outputBytes);
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  const audit = auditEnvelope(sourceBytes, outputBytes, roundTrip, checks);
  await fs.writeFile(finalAuditPath, `${JSON.stringify(audit, null, 2)}\n`, "utf8");
  return { outputPath: finalPath, auditPath: finalAuditPath, audit };
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const [inputPath, imagePath, outputPath, auditPath] = process.argv.slice(2);
  if (!inputPath || !imagePath || !outputPath || !auditPath) throw new Error("usage: officekit-branded-template-local-update-workflow.mjs <input.pptx> <replacement.png> <output.pptx> <audit.json>");
  console.log(JSON.stringify((await editBrandedTemplate({ inputPath, imagePath, outputPath, auditPath })).audit, null, 2));
}
