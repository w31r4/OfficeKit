#!/usr/bin/env node

import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { mkdir, readFile, stat, writeFile } from "node:fs/promises";
import path from "node:path";

import { PresentationFile } from "../src/presentation/index.mjs";
import { FileBlob } from "../src/shared/file-blob.mjs";
import { createLibreOfficeRenderer } from "../src/renderers/libreoffice.mjs";
import { createPopplerRenderer } from "../src/renderers/poppler.mjs";
import { SOURCES } from "./pptx-six-sample-import.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const DEFAULT_ASSETS_DIR = path.resolve("tmp/reference-pptx-downloads");
const DEFAULT_OUTPUT_DIR = path.resolve("tmp/presentation-six-sample-render");
const MAX_SOURCE_BYTES = 128 * 1024 * 1024;

export async function collectSixSampleRenderEvidence({
  assetsDir = DEFAULT_ASSETS_DIR,
  outputDir = DEFAULT_OUTPUT_DIR,
} = {}) {
  const sourceRoot = path.resolve(assetsDir);
  const renderRoot = path.resolve(outputDir);
  if (renderRoot === sourceRoot || renderRoot.startsWith(`${sourceRoot}${path.sep}`)) {
    throw new Error("Render output must not be inside the source directory.");
  }
  await mkdir(renderRoot, { recursive: true });
  const office = createLibreOfficeRenderer({ timeoutMs: 120_000 });
  const raster = createPopplerRenderer({ dpi: 96, timeoutMs: 120_000 });
  const sources = [];
  for (const source of SOURCES) {
    const sourcePath = path.join(sourceRoot, source.fileName);
    const bytes = await readSource(sourcePath, source.sha256);
    const original = await importPresentation(bytes);
    const target = firstPlacementObject(original);
    const textTarget = target ? undefined : firstTextRun(original);
    if (!target && !textTarget) throw new Error(`${source.id} has no bounded placement or text target.`);
    const targetPage = target?.slide || textTarget.slide;
    const editKind = target ? "placement" : "text";
    if (target) {
      const before = { ...target.object.position };
      target.object.setPosition({ left: before.left + 3, top: before.top + 3 });
    } else {
      const sourceText = textTarget.run.text;
      textTarget.shape.text.replace(sourceText, `${sourceText} OfficeKit`);
    }
    const edited = await PresentationFile.exportPptx(original);
    const baseline = await renderPages(bytes, original.slides.count, path.join(renderRoot, source.id, "source"), office, raster);
    const output = await renderPages(edited.bytes, original.slides.count, path.join(renderRoot, source.id, editKind), office, raster);
    const pageResults = baseline.map((page, index) => ({
      slide: page.slide,
      sourcePngSha256: page.sha256,
      editedPngSha256: output[index]?.sha256,
      changed: page.sha256 !== output[index]?.sha256,
    }));
    if (pageResults.some((page) => page.slide !== targetPage && page.changed)) {
      throw new Error(`${source.id} changed a non-target rendered page.`);
    }
    sources.push({
      id: source.id,
      sourceSha256: source.sha256,
      slideCount: baseline.length,
      renderer: { office: "LibreOffice", raster: "Poppler", dpi: 96 },
      ...(target
        ? { placementTarget: { id: target.object.id, slide: targetPage } }
        : { textTarget: { id: textTarget.shape.id, slide: targetPage } }),
      editKind,
      targetPageChanged: pageResults.find((page) => page.slide === targetPage)?.changed === true,
      nonTargetPagesPixelIdentical: pageResults.filter((page) => page.slide !== targetPage).every((page) => !page.changed),
      pages: pageResults,
    });
  }
  return {
    schema: "office-kit/pptx-six-sample-render-evidence/v1",
    sources,
    totals: {
      sources: sources.length,
      slides: sources.reduce((sum, source) => sum + source.slideCount, 0),
      nonTargetPagesPixelIdentical: sources.every((source) => source.nonTargetPagesPixelIdentical),
      targetPagesChanged: sources.filter((source) => source.targetPageChanged).length,
    },
  };
}

async function renderPages(bytes, slideCount, outputDir, office, raster) {
  await mkdir(outputDir, { recursive: true });
  const pdf = await office({
    input: new FileBlob(bytes, { type: PPTX_MIME }),
    inputType: PPTX_MIME,
    outputType: "application/pdf",
    format: "pdf",
    artifactKind: "presentation",
  });
  const pdfPath = path.join(outputDir, "render.pdf");
  await pdf.save(pdfPath);
  const nativePageCount = pdfPageCount(pdfPath);
  if (nativePageCount !== slideCount) throw new Error(`LibreOffice rendered ${nativePageCount} pages for a ${slideCount}-slide presentation.`);
  const renderedPages = [];
  for (let index = 0; index < slideCount; index += 1) {
    const png = await raster({
      input: pdf,
      inputType: "application/pdf",
      outputType: "image/png",
      format: "png",
      artifactKind: "presentation",
      pageIndex: index,
    });
    const imagePath = path.join(outputDir, `slide-${String(index + 1).padStart(2, "0")}.png`);
    await png.save(imagePath);
    renderedPages.push({ slide: index + 1, sha256: sha256(png.bytes) });
  }
  return renderedPages;
}

function firstPlacementObject(presentation) {
  for (const slide of presentation.slides.items) {
    const object = (slide.nativeObjects?.items || []).find((candidate) => candidate.placementCapability?.supported === true);
    if (object) return { slide: slide.index + 1, object };
  }
  return undefined;
}

function firstTextRun(presentation) {
  for (const slide of presentation.slides.items) {
    const shapes = [];
    const collect = (group) => {
      shapes.push(...(group.shapes?.items || []));
      for (const child of group.groups?.items || []) collect(child);
    };
    shapes.push(...(slide.shapes?.items || []));
    for (const group of slide.groups?.items || []) collect(group);
    for (const shape of shapes) {
      for (const paragraph of shape.text?.paragraphs || []) {
        const run = paragraph.runs?.find((candidate) => typeof candidate.text === "string" && candidate.text.trim().length >= 4);
        if (run) return { slide: slide.index + 1, shape, run };
      }
    }
  }
  return undefined;
}

async function readSource(filePath, expectedSha256) {
  const info = await stat(filePath);
  if (!info.isFile() || info.size < 1 || info.size > MAX_SOURCE_BYTES) throw new RangeError(`PPTX input is outside 1..${MAX_SOURCE_BYTES}: ${filePath}`);
  const bytes = await readFile(filePath);
  if (sha256(bytes) !== expectedSha256) throw new Error(`Source SHA-256 mismatch for ${filePath}.`);
  return bytes;
}

async function importPresentation(bytes) {
  return PresentationFile.importPptx(new FileBlob(bytes, { type: PPTX_MIME }));
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function pdfPageCount(pdfPath) {
  const result = spawnSync("pdfinfo", [pdfPath], { encoding: "utf8", shell: false });
  if (result.status !== 0) throw new Error(`pdfinfo failed for ${pdfPath}: ${result.stderr || result.stdout}`);
  const pages = Number(/^Pages:\s+(\d+)/mu.exec(result.stdout)?.[1]);
  if (!Number.isInteger(pages) || pages < 1) throw new Error(`pdfinfo did not report a valid page count for ${pdfPath}.`);
  return pages;
}

function parseArgs(argv) {
  const options = {};
  for (let index = 0; index < argv.length; index += 1) {
    const key = argv[index];
    if (key === "--assets-dir") options.assetsDir = argv[++index];
    else if (key === "--output-dir") options.outputDir = argv[++index];
    else throw new Error(`Unknown option ${key}.`);
  }
  return options;
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const options = parseArgs(process.argv.slice(2));
  collectSixSampleRenderEvidence(options)
    .then(async (evidence) => {
      const outputDir = path.resolve(options.outputDir || DEFAULT_OUTPUT_DIR);
      const evidencePath = path.join(outputDir, "evidence.json");
      await writeFile(evidencePath, `${JSON.stringify(evidence, null, 2)}\n`, { flag: "wx" });
      process.stdout.write(`${JSON.stringify({ ok: true, evidence: evidencePath, ...evidence.totals })}\n`);
    })
    .catch((error) => {
      process.stderr.write(`${error?.stack || error}\n`);
      process.exitCode = 2;
    });
}
