#!/usr/bin/env node

import { createHash } from "node:crypto";
import { readFile, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

import { FileBlob, PresentationFile } from "../src/index.mjs";
import { directPresentationChildren } from "../src/presentation/group-shapes.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const DEFAULT_MANIFEST = path.resolve("evals/presentation-six-sample-import/manifest.v1.json");
const VISIBLE_ROOTS = new Set(["sp", "pic", "graphicFrame", "cxnSp", "grpSp", "contentPart", "AlternateContent"]);
const METADATA_ROOTS = new Set(["nvGrpSpPr", "grpSpPr", "extLst"]);

export async function inspectSixSamplePptx({ assetsDir, manifestPath = DEFAULT_MANIFEST } = {}) {
  if (!assetsDir) throw new TypeError("assetsDir is required.");
  const manifestFile = path.resolve(manifestPath);
  const manifestBytes = await readFile(manifestFile);
  const manifest = JSON.parse(manifestBytes);
  if (manifest.schema !== "office-kit/presentation-six-sample-import-manifest/v1" || !Array.isArray(manifest.sources) || manifest.sources.length !== 6) {
    throw new Error("Expected the six-sample import manifest.");
  }
  const samples = [];
  for (const source of manifest.sources) {
    const filePath = path.join(path.resolve(assetsDir), source.fileName);
    const bytes = await boundedRead(filePath);
    const sourceSha256 = sha256(bytes);
    if (sourceSha256 !== source.sha256) throw new Error(`${source.id}: source SHA-256 mismatch.`);
    const raw = await rawShapeTree(bytes);
    const presentation = await PresentationFile.importPptx(new FileBlob(bytes, { type: PPTX_MIME, name: source.fileName }));
    const imported = parseNdjson(presentation.inspect({ kind: "importObject", maxChars: Infinity }).ndjson);
    if (presentation.slides.count !== raw.slides || imported.length !== raw.objects) {
      throw new Error(`${source.id}: source has ${raw.slides}/${raw.objects}, import has ${presentation.slides.count}/${imported.length}.`);
    }
    assertLocators(source.id, imported, sourceSha256);
    const noOp = await PresentationFile.exportPptx(presentation);
    if (!Buffer.from(noOp.bytes).equals(bytes)) throw new Error(`${source.id}: no-op export changed source bytes.`);
    const opaqueText = parseNdjson(presentation.inspect({ kind: "nativeObject", maxChars: Infinity }).ndjson)
      .filter((record) => typeof record.text === "string" && record.text.length > 0);
    samples.push({
      id: source.id,
      fileName: source.fileName,
      bytes: bytes.byteLength,
      sha256: sourceSha256,
      slides: presentation.slides.count,
      objects: imported.length,
      typed: imported.filter((record) => record.classification === "typed-editable").length,
      classifications: counts(imported.map((record) => record.classification)),
      objectKinds: counts(imported.map((record) => record.objectKind)),
      opaqueTextObjects: opaqueText.length,
      noOpExact: true,
    });
  }
  return {
    schema: "office-kit/presentation-six-sample-import-runtime-evidence/v1",
    measuredAt: new Date().toISOString().slice(0, 10),
    sourceManifestSha256: sha256(manifestBytes),
    samples,
    totals: {
      slides: samples.reduce((sum, sample) => sum + sample.slides, 0),
      objects: samples.reduce((sum, sample) => sum + sample.objects, 0),
      exactNoOps: samples.filter((sample) => sample.noOpExact).length,
    },
  };
}

async function rawShapeTree(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const slides = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide[1-9][0-9]*[.]xml$/u.test(name));
  let objects = 0;
  for (const name of slides) {
    const xml = await zip.file(name).async("text");
    const roots = directPresentationChildren(xml, "spTree");
    if (roots.some((root) => !METADATA_ROOTS.has(root.localName) && !VISIBLE_ROOTS.has(root.localName))) {
      throw new Error(`Unsupported direct shape-tree root in ${name}.`);
    }
    objects += roots.filter((root) => VISIBLE_ROOTS.has(root.localName)).length;
  }
  return { slides: slides.length, objects };
}

function assertLocators(sourceId, records, sourceSha256) {
  const locators = records.map((record) => `${record.sourceLocator?.slideId}:${record.sourceLocator?.shapeTreeIndex}`);
  if (locators.some((locator) => locator.startsWith("undefined:")) || new Set(locators).size !== locators.length) {
    throw new Error(`${sourceId}: import locators are missing or duplicated.`);
  }
  if (records.some((record) => record.sourceRevisionSha256 !== sourceSha256)) {
    throw new Error(`${sourceId}: import locator is not source-revision bound.`);
  }
}

async function boundedRead(filePath) {
  const info = await stat(filePath);
  if (!info.isFile() || info.size < 1 || info.size > 128 * 1024 * 1024) throw new RangeError(`Invalid PPTX source: ${filePath}`);
  const bytes = await readFile(filePath);
  if (bytes.byteLength !== info.size) throw new Error(`PPTX source changed while reading: ${filePath}`);
  return bytes;
}

function counts(values) {
  const result = {};
  for (const value of values) result[value] = (result[value] || 0) + 1;
  return Object.fromEntries(Object.entries(result).sort(([a], [b]) => a.localeCompare(b)));
}

function parseNdjson(value) {
  return String(value).trim().split("\n").filter(Boolean).map((line) => JSON.parse(line));
}

function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

async function main() {
  const argv = process.argv.slice(2);
  const assetsIndex = argv.indexOf("--assets-dir");
  const outputIndex = argv.indexOf("--output");
  const manifestIndex = argv.indexOf("--manifest");
  const assetsDir = assetsIndex >= 0 ? argv[assetsIndex + 1] : undefined;
  const output = outputIndex >= 0 ? argv[outputIndex + 1] : undefined;
  const manifestPath = manifestIndex >= 0 ? argv[manifestIndex + 1] : DEFAULT_MANIFEST;
  if (!assetsDir || !output) throw new Error("Usage: pptx-six-sample-import.mjs --assets-dir <dir> --output <file> [--manifest <file>]");
  const evidence = await inspectSixSamplePptx({ assetsDir, manifestPath });
  await writeFile(path.resolve(output), `${JSON.stringify(evidence, null, 2)}\n`, { flag: "wx" });
  process.stdout.write(`${JSON.stringify({ ok: true, ...evidence.totals })}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) main().catch((error) => {
  process.stderr.write(`${error?.stack || error}\n`);
  process.exitCode = 2;
});
