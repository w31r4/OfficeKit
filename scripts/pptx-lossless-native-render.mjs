#!/usr/bin/env node

import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

import { FileBlob, PresentationFile, visualQaArtifact } from "../src/index.mjs";
import { createLibreOfficeRenderer } from "../src/renderers/libreoffice.mjs";
import { createPopplerRenderer } from "../src/renderers/poppler.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const REPOSITORY_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DEFAULT_MANIFEST = path.join(REPOSITORY_ROOT, "evals/pptx-lossless/manifest.v1.json");
const DEFAULT_EVIDENCE = path.join(REPOSITORY_ROOT, "evals/pptx-lossless/evidence.v1.json");

const options = parseArgs(process.argv.slice(2));
if (!options.assetsDir || !options.editsDir || !options.outputDir) {
  fail("Usage: pptx-lossless-native-render.mjs --assets-dir <dir> --edits-dir <dir> --output-dir <dir> [--manifest <path>] [--evidence <path>]");
}

const assetsDir = path.resolve(options.assetsDir);
const editsDir = path.resolve(options.editsDir);
const outputDir = path.resolve(options.outputDir);
const manifest = JSON.parse(await readFile(path.resolve(options.manifest || DEFAULT_MANIFEST), "utf8"));
const expectedEvidence = JSON.parse(await readFile(path.resolve(options.evidence || DEFAULT_EVIDENCE), "utf8"));
assertSeparateRoot(assetsDir, outputDir);
assertSeparateRoot(editsDir, outputDir);
await mkdir(outputDir, { recursive: true });

const libreOffice = createLibreOfficeRenderer({ timeoutMs: 120_000 });
const poppler = createPopplerRenderer({ dpi: 144, timeoutMs: 120_000 });
const sources = [];
for (const source of manifest.sources) {
  const sourcePath = resolveSourcePath(source, assetsDir);
  const sourceBytes = await readFile(sourcePath);
  if (sha256(sourceBytes) !== source.sha256) fail(`Source hash mismatch for ${source.id}.`);
  const sourceBlob = new FileBlob(sourceBytes, { type: PPTX_MIME, name: path.basename(sourcePath) });
  const presentation = await PresentationFile.importPptx(sourceBlob);
  const sourceRoot = path.join(outputDir, source.id, "source");
  const baselinePages = await renderNativePages(sourceBlob, presentation.slides.count, sourceRoot, libreOffice, poppler);
  const sourceEvidence = expectedEvidence.sources.find((candidate) => candidate.id === source.id);
  if (!sourceEvidence?.noOpByteIdentical) fail(`Versioned no-op evidence is missing for ${source.id}.`);
  const targets = [];
  for (const target of source.targets) {
    const expected = sourceEvidence.targets.find((candidate) => candidate.id === target.id);
    if (!expected) fail(`Versioned target evidence is missing for ${source.id}/${target.id}.`);
    const editPath = path.join(editsDir, `${source.id}-${target.id}-1.pptx`);
    const editBytes = await readFile(editPath);
    if (sha256(editBytes) !== expected.outputSha256) fail(`Unexpected benchmark output hash for ${source.id}/${target.id}.`);
    const editBlob = new FileBlob(editBytes, { type: PPTX_MIME, name: path.basename(editPath) });
    const editRoot = path.join(outputDir, source.id, target.id);
    const outputPages = await renderNativePages(editBlob, presentation.slides.count, editRoot, libreOffice, poppler);
    const changedSlide = targetSlideNumber(target);
    const pages = [];
    for (let index = 0; index < baselinePages.length; index += 1) {
      const baseline = baselinePages[index];
      const actual = outputPages[index];
      const qa = await visualQaArtifact({ render: () => actual.blob }, {
        baseline: baseline.blob,
        pixelDiff: true,
        allowPixelChange: true,
        diffImage: false,
        minBytes: 100,
        maxChars: 2_000,
      });
      const changed = qa.summary.pixelDiff?.changed;
      if (changed == null || qa.summary.pixelDiff?.skipped) fail(`Native pixel comparison did not complete for ${source.id}/${target.id}/slide-${index + 1}.`);
      if (index + 1 !== changedSlide && changed) fail(`Non-target slide changed after ${source.id}/${target.id}: slide ${index + 1}.`);
      pages.push({ slide: index + 1, sourceHash: baseline.hash, outputHash: actual.hash, changed, pixelDiff: qa.summary.pixelDiff });
    }
    const targetPage = pages[changedSlide - 1];
    targets.push({
      id: target.id,
      outputSha256: expected.outputSha256,
      changedSlide,
      pages,
      nonTargetPagesPixelIdentical: true,
      targetPageVisualState: targetPage?.changed ? "changed" : "unchanged-in-libreoffice",
    });
  }
  sources.push({
    id: source.id,
    sourceSha256: source.sha256,
    nativePageCount: presentation.slides.count,
    baselinePages: baselinePages.map(({ slide, hash }) => ({ slide, hash })),
    targets,
  });
}

const result = {
  schema: "office-kit/pptx-lossless-native-visual-evidence/v1",
  renderer: { office: "LibreOffice", raster: "Poppler", dpi: 144 },
  sources,
};
const evidencePath = path.join(outputDir, "evidence.json");
await writeFile(evidencePath, `${JSON.stringify(result, null, 2)}\n`, { flag: "wx" });
console.log(JSON.stringify({ ok: true, evidence: evidencePath, sources: sources.length }));

function parseArgs(argv) {
  const parsed = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) fail(`Unexpected argument: ${token}`);
    const key = token.slice(2).replace(/-([a-z])/gu, (_, letter) => letter.toUpperCase());
    const value = argv[++index];
    if (!value || value.startsWith("--")) fail(`Missing value for --${key}.`);
    parsed[key] = value;
  }
  return parsed;
}

function resolveSourcePath(source, externalAssetsDir) {
  return source.sourceKind === "repository-supplemental"
    ? path.resolve(REPOSITORY_ROOT, source.fileName)
    : path.resolve(externalAssetsDir, source.fileName);
}

async function renderNativePages(pptx, slideCount, root, officeRenderer, rasterRenderer) {
  await mkdir(root, { recursive: true });
  const pdf = await officeRenderer({
    input: pptx,
    inputType: PPTX_MIME,
    outputType: "application/pdf",
    format: "pdf",
    artifactKind: "presentation",
  });
  const pdfPath = path.join(root, "native.pdf");
  await pdf.save(pdfPath);
  const nativePageCount = pdfPageCount(pdfPath);
  if (nativePageCount !== slideCount) fail(`Native render produced ${nativePageCount} pages for a ${slideCount}-slide presentation.`);
  const pages = [];
  for (let pageIndex = 0; pageIndex < slideCount; pageIndex += 1) {
    const blob = await rasterRenderer({
      input: pdf,
      inputType: "application/pdf",
      outputType: "image/png",
      format: "png",
      artifactKind: "presentation",
      pageIndex,
    });
    const slide = pageIndex + 1;
    await blob.save(path.join(root, `slide-${String(slide).padStart(2, "0")}.png`));
    pages.push({ slide, blob, hash: sha256(blob.bytes) });
  }
  return pages;
}

function pdfPageCount(pdfPath) {
  const result = spawnSync("pdfinfo", [pdfPath], { encoding: "utf8", shell: false });
  if (result.status !== 0) fail(`pdfinfo failed for ${pdfPath}: ${result.stderr || result.stdout}`);
  const pages = Number(/^Pages:\s+(\d+)/mu.exec(result.stdout)?.[1]);
  if (!Number.isInteger(pages) || pages < 1) fail(`pdfinfo did not report a valid page count for ${pdfPath}.`);
  return pages;
}

function targetSlideNumber(target) {
  const value = /^presentation\/slide\/(\d+)\/element\//u.exec(String(target.nodeId || ""))?.[1];
  const slide = Number(value);
  if (!Number.isInteger(slide) || slide < 1) fail(`Target ${target.id} does not declare a slide-scoped node ID.`);
  return slide;
}

function assertSeparateRoot(input, output) {
  if (input === output || output.startsWith(`${input}${path.sep}`)) fail("Output directory must be outside every input directory.");
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function fail(message) {
  throw new Error(message);
}
