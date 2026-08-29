import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import sharp from "sharp";

import { Presentation, PresentationFile } from "../src/presentation/index.mjs";

const PROBE_TIMEOUT_MS = 5_000;
const RENDER_TIMEOUT_MS = 120_000;

function commandAvailable(command, args = ["--version"]) {
  const result = spawnSync(command, args, {
    encoding: "utf8",
    timeout: PROBE_TIMEOUT_MS,
    killSignal: "SIGKILL",
    maxBuffer: 64 * 1024,
  });
  return !result.error && result.status === 0;
}

function run(command, args) {
  const result = spawnSync(command, args, {
    encoding: "utf8",
    timeout: RENDER_TIMEOUT_MS,
    killSignal: "SIGKILL",
    maxBuffer: 1024 * 1024,
  });
  assert.equal(result.error, undefined, `${command} could not start: ${result.error?.message || "unknown error"}`);
  assert.equal(result.status, 0, `${command} failed with ${result.status}: ${String(result.stderr || result.stdout).trim()}`);
  return result;
}

function probePositions(bboxHtml) {
  return [...bboxHtml.matchAll(/<page\b[^>]*>([\s\S]*?)<\/page>/g)].map((page, index) => {
    const word = /<word\b([^>]*)>RECTPROBE<\/word>/.exec(page[1]);
    assert.ok(word, `LibreOffice/Poppler page ${index + 1} must expose RECTPROBE as selectable text.`);
    const x = Number(/\bxMin="([^"]+)"/.exec(word[1])?.[1]);
    const y = Number(/\byMin="([^"]+)"/.exec(word[1])?.[1]);
    assert.ok(Number.isFinite(x) && Number.isFinite(y), `Page ${index + 1} needs finite RECTPROBE coordinates.`);
    return { x, y };
  });
}

async function greenCentroid(filePath) {
  const { data, info } = await sharp(filePath).removeAlpha().raw().toBuffer({ resolveWithObject: true });
  let count = 0;
  let sumX = 0;
  for (let y = 0; y < info.height; y += 1) {
    for (let x = 0; x < info.width; x += 1) {
      const offset = (y * info.width + x) * info.channels;
      const [red, green, blue] = data.subarray(offset, offset + 3);
      if (green > 90 && green > red * 1.35 && green > blue * 1.2) {
        count += 1;
        sumX += x;
      }
    }
  }
  assert.ok(count > 2_000, `${filePath} must contain a substantial green formula shape.`);
  return sumX / count;
}

if (!commandAvailable("soffice") || !commandAvailable("pdftotext", ["-v"]) || !commandAvailable("pdftoppm", ["-v"])) {
  console.log("presentation custom geometry native render skipped (LibreOffice/Poppler unavailable)");
  process.exit(0);
}

const work = await fs.mkdtemp(path.join(os.tmpdir(), "officekit-custom-text-rectangle-"));
try {
  const presentation = Presentation.create({ slideSize: { width: 640, height: 360 } });
  const textRectangleCases = [
    { name: "Default text bounds" },
    { name: "Inset text bounds", textRectangle: { left: 180, top: 90, right: 520, bottom: 240 } },
    {
      name: "Guide text bounds",
      textRectangle: { left: "rectLeft", top: "rectTop", right: "rectRight", bottom: "rectBottom" },
      customGuides: [
        { name: "rectLeft", formula: "*/ w 180 560" },
        { name: "rectTop", formula: "*/ h 90 280" },
        { name: "rectRight", formula: "*/ w 520 560" },
        { name: "rectBottom", formula: "*/ h 240 280" },
      ],
    },
  ];
  for (const { name, textRectangle, customGuides } of textRectangleCases) {
    const slide = presentation.slides.add({ name });
    slide.shapes.add({
      name: `probe-${slide.index + 1}`,
      geometry: "custom",
      position: { left: 40, top: 40, width: 560, height: 280 },
      fill: "#DBEAFE",
      line: { fill: "#2563EB", width: 2 },
      text: "RECTPROBE",
      textRectangle,
      customGuides,
      textStyle: { fontFamily: "Liberation Sans", fontSize: 24, color: "#0F172A" },
      textBodyProperties: { anchor: "top", insets: { left: 0, top: 0, right: 0, bottom: 0 } },
      customPaths: [{
        width: 100,
        height: 100,
        commands: [
          { moveTo: { x: 50, y: 0 } },
          { lineTo: { x: 100, y: 50 } },
          { lineTo: { x: 50, y: 100 } },
          { lineTo: { x: 0, y: 50 } },
          { close: {} },
        ],
      }],
    });
  }
  const formulaSlides = [];
  for (const [name, adjustment] of [["Left formula apex", 25_000], ["Right formula apex", 75_000]]) {
    const slide = presentation.slides.add({ name });
    formulaSlides.push(slide);
    const position = { left: 40, top: 40, width: 560, height: 280 };
    slide.shapes.add({
      name: `formula-probe-${slide.index + 1}`,
      geometry: "custom",
      position,
      fill: "#16A34A",
      line: { fill: "transparent", width: 0 },
      customAdjustments: [{ name: "adjX", formula: `val ${adjustment}` }],
      customGuides: [{ name: "apexX", formula: "*/ w adjX 100000" }],
      customPaths: [{
        commands: [
          { moveTo: { x: "l", y: "b" } },
          { lineTo: { x: "apexX", y: "t" } },
          { lineTo: { x: "r", y: "b" } },
          { close: {} },
        ],
      }],
    });
  }

  const modelFormulaCentroids = [];
  for (const [index, slide] of formulaSlides.entries()) {
    const modelPath = path.join(work, `model-formula-${index + 3}.png`);
    const svg = await (await slide.export()).text();
    await sharp(Buffer.from(svg)).png().toFile(modelPath);
    modelFormulaCentroids.push(await greenCentroid(modelPath));
  }
  assert.ok(modelFormulaCentroids[1] - modelFormulaCentroids[0] > 35, `Guide adjustment must move the model-rendered formula-path centroid right: ${JSON.stringify(modelFormulaCentroids)}`);

  const pptxPath = path.join(work, "custom-text-rectangle.pptx");
  await (await PresentationFile.exportPptx(presentation)).save(pptxPath);
  const profilePath = path.join(work, "libreoffice-profile");
  run("soffice", [
    `-env:UserInstallation=${pathToFileURL(profilePath).href}`,
    "--headless",
    "--convert-to", "pdf:impress_pdf_Export",
    "--outdir", work,
    pptxPath,
  ]);
  const pdfPath = path.join(work, "custom-text-rectangle.pdf");
  const bboxPath = path.join(work, "custom-text-rectangle.html");
  assert.ok((await fs.stat(pdfPath)).size > 0, "LibreOffice must produce a non-empty PDF.");
  run("pdftotext", ["-f", "1", "-l", "3", "-bbox-layout", pdfPath, bboxPath]);
  const positions = probePositions(await fs.readFile(bboxPath, "utf8"));
  assert.equal(positions.length, 3, "Native PDF must retain all text-rectangle slides.");
  assert.ok(positions[1].x - positions[0].x > 80, `Inset text must move right in native output: ${JSON.stringify(positions)}`);
  assert.ok(positions[1].y - positions[0].y > 35, `Inset text must move down in native output: ${JSON.stringify(positions)}`);
  assert.ok(Math.abs(positions[2].x - positions[1].x) < 2, `Guide and numeric text rectangles must align horizontally: ${JSON.stringify(positions)}`);
  assert.ok(Math.abs(positions[2].y - positions[1].y) < 2, `Guide and numeric text rectangles must align vertically: ${JSON.stringify(positions)}`);
  run("pdftoppm", ["-f", "4", "-l", "5", "-r", "72", "-png", pdfPath, path.join(work, "formula")]);
  const nativeFormulaCentroids = [
    await greenCentroid(path.join(work, "formula-4.png")),
    await greenCentroid(path.join(work, "formula-5.png")),
  ];
  assert.ok(nativeFormulaCentroids[1] - nativeFormulaCentroids[0] > 35, `Guide adjustment must move the native-rendered formula-path centroid right: ${JSON.stringify(nativeFormulaCentroids)}`);
  console.log(`presentation custom geometry render ok ${JSON.stringify({ positions, modelFormulaCentroids, nativeFormulaCentroids })}`);
} finally {
  if (process.env.OFFICEKIT_KEEP_QA === "1") {
    console.error(`presentation custom geometry QA retained at ${work}`);
  } else {
    await fs.rm(work, { recursive: true, force: true });
  }
}
