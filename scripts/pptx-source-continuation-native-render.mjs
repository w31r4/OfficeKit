#!/usr/bin/env node

import { createHash } from "node:crypto";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { FileBlob, PresentationFile, visualQaArtifact } from "../src/index.mjs";
import { createLibreOfficeRenderer } from "../src/renderers/libreoffice.mjs";
import { createPopplerRenderer } from "../src/renderers/poppler.mjs";
import { applyContinuation, verifyContinuation } from "./pptx-source-continuation-benchmark.mjs";
import { SOURCES } from "./pptx-source-reuse-benchmark.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const EVIDENCE_SCHEMA = "office-kit/pptx-source-continuation-native-evidence/v2";

export async function runSourceContinuationNativeRender(assetsDir, outputDir) {
  const office = createLibreOfficeRenderer({ timeoutMs: 180_000 });
  const poppler = createPopplerRenderer({ dpi: 120, timeoutMs: 180_000 });
  const results = [];
  for (const source of SOURCES) {
    const sourcePath = path.join(assetsDir, source.fileName);
    const sourceBytes = await readFile(sourcePath);
    const sourceBlob = new FileBlob(sourceBytes, { type: PPTX_MIME, name: source.fileName });
    const presentation = await PresentationFile.importPptx(sourceBlob);
    const sourceSlideCount = presentation.slides.count;
    const origin = presentation.slides.items[source.slide - 1];
    if (!origin.cloneCapability.supported) throw new Error(`${source.id} clone capability is not supported: ${origin.cloneCapability.blockedReason}`);
    const pendingClone = origin.duplicate();
    // Keep all original slide numbers and page positions stable; the derived
    // page is appended so the render oracle can require exact non-target pixels.
    pendingClone.moveTo(sourceSlideCount);
    const cloneOutput = await PresentationFile.exportPptx(presentation);
    const continued = await PresentationFile.importPptx(cloneOutput.bytes);
    const clone = continued.slides.items[sourceSlideCount];
    const target = applyContinuation(clone, "bounded-overlay");
    const output = await PresentationFile.exportPptx(continued);
    const verified = await PresentationFile.importPptx(output.bytes);
    const verifiedTarget = verifyContinuation(verified.slides.items[sourceSlideCount], target);
    const sourcePages = await renderNativePages(sourceBlob, sourceSlideCount, outputDir, `${source.id}-source`, office, poppler);
    const clonedInsertedPage = await renderNativePage(new FileBlob(cloneOutput.bytes, { type: PPTX_MIME, name: `${source.id}-clone.pptx` }), sourceSlideCount, outputDir, `${source.id}-clone`, office, poppler);
    const outputPages = await renderNativePages(new FileBlob(output.bytes, { type: PPTX_MIME, name: `${source.id}-continuation.pptx` }), verified.slides.count, outputDir, `${source.id}-continuation`, office, poppler);
    if (outputPages.length !== sourcePages.length + 1) throw new Error(`${source.id} native output page count ${outputPages.length} does not equal source count + 1.`);
    const pages = [];
    for (let index = 0; index < sourcePages.length; index += 1) {
      const outputIndex = index;
      const baseline = sourcePages[index];
      const actual = outputPages[outputIndex];
      const qa = await visualQaArtifact({ render: () => actual.blob }, {
        baseline: baseline.blob,
        pixelDiff: true,
        allowPixelChange: true,
        diffImage: false,
        minBytes: 100,
        maxChars: 2_000,
      });
      if (qa.summary.pixelDiff?.skipped || qa.summary.pixelDiff?.changed !== false) {
        throw new Error(`${source.id} non-target slide ${index + 1} changed in native rendering: ${JSON.stringify(qa.summary.pixelDiff)}`);
      }
      pages.push({ sourceSlide: index + 1, outputSlide: outputIndex + 1, sourceHash: baseline.hash, outputHash: actual.hash, pixelIdentical: true });
    }
    const inserted = outputPages[sourceSlideCount];
    if (!inserted?.blob?.bytes?.length) throw new Error(`${source.id} inserted continuation page did not render.`);
    const insertedQa = await visualQaArtifact({ render: () => inserted.blob }, {
      baseline: clonedInsertedPage.blob,
      pixelDiff: true,
      allowPixelChange: true,
      diffImage: false,
      minBytes: 100,
      maxChars: 2_000,
    });
    const insertedPixelDiff = insertedQa.summary.pixelDiff;
    if (insertedPixelDiff?.skipped || insertedPixelDiff?.changed !== true) {
      throw new Error(`${source.id} continuation overlay was not visible in native rendering: ${JSON.stringify(insertedPixelDiff)}`);
    }
    results.push({
      id: source.id,
      sourceSha256: sha256(sourceBytes),
      sourceSlideCount: sourcePages.length,
      outputSlideCount: outputPages.length,
      insertedSlide: sourceSlideCount + 1,
      continuationKind: "bounded-overlay",
      target: verifiedTarget,
      nonTargetPagesPixelIdentical: true,
      insertedPageRendered: true,
      insertedPageChangedFromClone: true,
      insertedPageChange: {
        cloneHash: clonedInsertedPage.hash,
        outputHash: inserted.hash,
        differentPixels: insertedPixelDiff.differentPixels,
        mismatchRatio: insertedPixelDiff.mismatchRatio,
      },
      pages,
    });
  }
  return { schema: EVIDENCE_SCHEMA, renderer: { office: "LibreOffice", raster: "Poppler", dpi: 120 }, sources: results };
}

async function renderNativePages(blob, slideCount, outputDir, prefix, office, poppler) {
  const root = await mkdtemp(path.join(outputDir, `${prefix}-`));
  const pdf = await office({ input: blob, inputType: PPTX_MIME, outputType: "application/pdf", format: "pdf", artifactKind: "presentation" });
  const pdfPath = path.join(root, "render.pdf");
  await pdf.save(pdfPath);
  const pages = [];
  for (let pageIndex = 0; pageIndex < slideCount; pageIndex += 1) {
    const image = await poppler({ input: pdf, inputType: "application/pdf", outputType: "image/png", format: "png", artifactKind: "presentation", pageIndex });
    pages.push({ slide: pageIndex + 1, blob: image, hash: sha256(image.bytes) });
  }
  return pages;
}

async function renderNativePage(blob, pageIndex, outputDir, prefix, office, poppler) {
  const root = await mkdtemp(path.join(outputDir, `${prefix}-`));
  const pdf = await office({ input: blob, inputType: PPTX_MIME, outputType: "application/pdf", format: "pdf", artifactKind: "presentation" });
  const pdfPath = path.join(root, "render.pdf");
  await pdf.save(pdfPath);
  const image = await poppler({ input: pdf, inputType: "application/pdf", outputType: "image/png", format: "png", artifactKind: "presentation", pageIndex });
  return { slide: pageIndex + 1, blob: image, hash: sha256(image.bytes) };
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function parseArgs(argv) {
  let assetsDir;
  let output;
  let force = false;
  for (let index = 0; index < argv.length; index += 1) {
    const flag = argv[index];
    if (flag === "--assets-dir") assetsDir = argv[++index];
    else if (flag === "--output") output = argv[++index];
    else if (flag === "--force") force = true;
    else throw new Error(`Unknown option ${flag}.`);
  }
  if (!assetsDir || !output) throw new Error("Usage: pptx-source-continuation-native-render.mjs --assets-dir <dir> --output <evidence.json> [--force]");
  return { assetsDir: path.resolve(assetsDir), output: path.resolve(output), force };
}

async function main() {
  const { assetsDir, output, force } = parseArgs(process.argv.slice(2));
  const outputDir = await mkdtemp(path.join(os.tmpdir(), "officekit-pptx-source-continuation-native-"));
  const evidence = await runSourceContinuationNativeRender(assetsDir, outputDir);
  await writeFile(output, `${JSON.stringify(evidence, null, 2)}\n`, { flag: force ? "w" : "wx" });
  process.stdout.write(`${JSON.stringify({ ok: true, output, sources: evidence.sources.length })}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
