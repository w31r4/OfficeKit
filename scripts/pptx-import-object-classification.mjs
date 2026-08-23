#!/usr/bin/env node

import { createHash } from "node:crypto";
import { mkdir, readFile, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

import { FileBlob, PresentationFile } from "../src/index.mjs";
import { directPresentationChildren } from "../src/presentation/group-shapes.mjs";

const SCHEMA = "office-kit/pptx-import-object-classification-evidence/v1";
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const DEFAULT_MANIFEST = path.resolve("evals/pptx-lossless/manifest.v1.json");
const MAX_SOURCE_BYTES = 128 * 1024 * 1024;
const MAX_SLIDES = 10_000;
const MAX_VISIBLE_OBJECTS = 100_000;
const METADATA_ROOTS = new Set(["nvGrpSpPr", "grpSpPr", "extLst"]);
const VISIBLE_ROOTS = new Set(["sp", "pic", "graphicFrame", "cxnSp", "grpSp", "contentPart", "AlternateContent"]);
const ROOT_KIND_COMPATIBILITY = Object.freeze({
  sp: new Set(["shape"]),
  pic: new Set(["image", "picture"]),
  graphicFrame: new Set(["table", "chart", "graphicFrame", "diagram", "oleObject"]),
  cxnSp: new Set(["connector"]),
  grpSp: new Set(["group"]),
  contentPart: new Set(["contentPart"]),
});

export async function buildImportObjectClassificationEvidence({ assetsDir, manifestPath = DEFAULT_MANIFEST } = {}) {
  if (!assetsDir) throw new TypeError("assetsDir is required.");
  const root = path.resolve(assetsDir);
  const manifestFile = path.resolve(manifestPath);
  const manifestBytes = await readFile(manifestFile);
  const manifest = JSON.parse(manifestBytes);
  if (manifest.schema !== "office-kit/pptx-lossless-benchmark/v1" || !Array.isArray(manifest.sources)) {
    throw new Error("PPTX lossless benchmark manifest has an unsupported schema.");
  }
  const definitions = manifest.sources.filter((source) => source.sourceKind !== "repository-supplemental");
  if (definitions.length !== 3) throw new Error(`Expected three external PPTX benchmark sources, received ${definitions.length}.`);
  const sources = [];
  for (const definition of definitions) {
    const sourcePath = path.join(root, definition.fileName);
    const bytes = await boundedRead(sourcePath);
    const digest = sha256(bytes);
    if (digest !== definition.sha256) throw new Error(`Source hash mismatch for ${definition.id}: ${digest}.`);
    sources.push(await classifySource(definition, bytes));
  }
  return {
    schema: SCHEMA,
    manifestSha256: sha256(manifestBytes),
    oracle: {
      rawSourceIndependent: true,
      directShapeTreeChildren: true,
      runtimeSelfReportedCountInsufficient: true,
      exactNoOpBytes: true,
    },
    totals: {
      sources: sources.length,
      slides: sum(sources, "slideCount"),
      visibleTopLevelObjects: sum(sources, "visibleTopLevelObjects"),
      classifiedTopLevelObjects: sum(sources, "classifiedTopLevelObjects"),
    },
    sources,
  };
}

async function classifySource(definition, bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const slideNames = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide[1-9][0-9]*[.]xml$/u.test(name))
    .sort((left, right) => slideOrdinal(left) - slideOrdinal(right));
  if (!slideNames.length || slideNames.length > MAX_SLIDES) throw new Error(`${definition.id} has an invalid slide count.`);

  const rawSlides = [];
  let rawTotal = 0;
  for (const [index, slideName] of slideNames.entries()) {
    const xml = await zip.file(slideName).async("text");
    const roots = directPresentationChildren(xml, "spTree");
    const unknown = roots.filter((root) => !METADATA_ROOTS.has(root.localName) && !VISIBLE_ROOTS.has(root.localName));
    if (unknown.length) {
      throw new Error(`${definition.id} slide ${index + 1} has unsupported direct shape-tree roots: ${[...new Set(unknown.map((root) => root.localName))].join(", ")}.`);
    }
    const visible = roots.filter((root) => VISIBLE_ROOTS.has(root.localName));
    rawTotal += visible.length;
    if (rawTotal > MAX_VISIBLE_OBJECTS) throw new Error(`${definition.id} exceeds the visible-object budget.`);
    rawSlides.push({
      slide: index + 1,
      slideId: `presentation/slide/${index + 1}`,
      roots: visible.map((root, shapeTreeIndex) => ({ shapeTreeIndex, rootKind: root.localName })),
    });
  }

  const presentation = await PresentationFile.importPptx(new FileBlob(bytes, { type: PPTX_MIME }));
  const records = parseNdjson(presentation.inspect({ kind: "importObject", maxChars: Infinity }).ndjson);
  if (presentation.slides.count !== rawSlides.length) {
    throw new Error(`${definition.id} imported ${presentation.slides.count} slides from ${rawSlides.length} source slides.`);
  }
  const classifiedBySlide = new Map();
  for (const record of records) {
    const values = classifiedBySlide.get(record.slide) || [];
    values.push(record);
    classifiedBySlide.set(record.slide, values);
  }

  const slides = [];
  for (const rawSlide of rawSlides) {
    const classified = [...(classifiedBySlide.get(rawSlide.slide) || [])]
      .sort((left, right) => left.sourceLocator.shapeTreeIndex - right.sourceLocator.shapeTreeIndex);
    if (classified.length !== rawSlide.roots.length) {
      throw new Error(`${definition.id} slide ${rawSlide.slide} classified ${classified.length} of ${rawSlide.roots.length} visible source objects.`);
    }
    for (const [index, root] of rawSlide.roots.entries()) {
      const record = classified[index];
      if (record.sourceLocator.slideId !== rawSlide.slideId || record.sourceLocator.shapeTreeIndex !== root.shapeTreeIndex) {
        throw new Error(`${definition.id} slide ${rawSlide.slide} is missing source object index ${root.shapeTreeIndex}.`);
      }
      assertCompatibleRoot(definition.id, rawSlide.slide, root, record);
    }
    slides.push({
      slide: rawSlide.slide,
      visibleTopLevelObjects: rawSlide.roots.length,
      classifiedTopLevelObjects: classified.length,
      rawRootKinds: counts(rawSlide.roots.map((root) => root.rootKind)),
      classifications: counts(classified.map((record) => record.classification)),
    });
  }
  if (records.length !== rawTotal) throw new Error(`${definition.id} classified ${records.length} of ${rawTotal} visible source objects.`);
  if (new Set(records.map((record) => record.targetId)).size !== records.length) throw new Error(`${definition.id} has duplicate classified target IDs.`);
  if (new Set(records.map((record) => `${record.sourceLocator.slideId}:${record.sourceLocator.shapeTreeIndex}`)).size !== records.length) {
    throw new Error(`${definition.id} has duplicate classified source locators.`);
  }
  if (records.some((record) => record.sourceRevisionSha256 !== definition.sha256)) {
    throw new Error(`${definition.id} classification records are not bound to the immutable source revision.`);
  }
  const serialized = JSON.stringify(records);
  if (/rawXml|partPath|relationshipId|sourcePackage|<p:/u.test(serialized)) {
    throw new Error(`${definition.id} classification records expose a forbidden package selector.`);
  }
  const noOp = await PresentationFile.exportPptx(presentation);
  if (!Buffer.from(noOp.bytes).equals(bytes)) throw new Error(`${definition.id} classification inspection changed no-op export bytes.`);
  return {
    id: definition.id,
    fileName: definition.fileName,
    bytes: bytes.byteLength,
    sha256: definition.sha256,
    slideCount: rawSlides.length,
    visibleTopLevelObjects: rawTotal,
    classifiedTopLevelObjects: records.length,
    complete: true,
    rawRootKinds: counts(rawSlides.flatMap((slide) => slide.roots.map((root) => root.rootKind))),
    objectKinds: counts(records.map((record) => record.objectKind)),
    classifications: counts(records.map((record) => record.classification)),
    typedOperations: counts(records.flatMap((record) => record.typedOperations)),
    nativeLeafKinds: counts(records.flatMap((record) => record.nativeLeafKinds)),
    slidesEvidenceSha256: sha256(Buffer.from(JSON.stringify(slides))),
    slides,
    noOpByteIdentical: true,
  };
}

function assertCompatibleRoot(sourceId, slide, root, record) {
  const accepted = ROOT_KIND_COMPATIBILITY[root.rootKind];
  if (accepted && !accepted.has(record.objectKind)) {
    throw new Error(`${sourceId} slide ${slide} source ${root.rootKind} at index ${root.shapeTreeIndex} became ${record.objectKind}.`);
  }
}

async function boundedRead(filePath) {
  const info = await stat(filePath);
  if (!info.isFile() || info.size < 1 || info.size > MAX_SOURCE_BYTES) {
    throw new RangeError(`PPTX benchmark source must be a regular file between 1 and ${MAX_SOURCE_BYTES} bytes.`);
  }
  const bytes = await readFile(filePath);
  if (bytes.byteLength !== info.size) throw new Error(`PPTX benchmark source changed while reading: ${path.basename(filePath)}.`);
  return bytes;
}

function slideOrdinal(name) {
  return Number(/slide([1-9][0-9]*)[.]xml$/u.exec(name)?.[1]);
}

function parseNdjson(value) {
  return String(value).trim().split("\n").filter(Boolean).map((line) => JSON.parse(line));
}

function counts(values) {
  const result = {};
  for (const value of values) result[value] = (result[value] || 0) + 1;
  return Object.fromEntries(Object.entries(result).sort(([left], [right]) => left.localeCompare(right)));
}

function sum(items, key) {
  return items.reduce((total, item) => total + Number(item[key] || 0), 0);
}

function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

function parseArgs(argv) {
  let assetsDir;
  let manifestPath = DEFAULT_MANIFEST;
  let output;
  let force = false;
  for (let index = 0; index < argv.length; index += 1) {
    const flag = argv[index];
    if (flag === "--assets-dir") assetsDir = argv[++index];
    else if (flag === "--manifest") manifestPath = argv[++index];
    else if (flag === "--output") output = argv[++index];
    else if (flag === "--force") force = true;
    else throw new Error(`Unknown option ${flag}.`);
  }
  if (!assetsDir || !output) {
    throw new Error("Usage: pptx-import-object-classification.mjs --assets-dir <dir> --output <evidence.json> [--manifest <path>] [--force]");
  }
  return { assetsDir, manifestPath, output, force };
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const evidence = await buildImportObjectClassificationEvidence(options);
  const output = path.resolve(options.output);
  await mkdir(path.dirname(output), { recursive: true });
  await writeFile(output, `${JSON.stringify(evidence, null, 2)}\n`, { flag: options.force ? "w" : "wx" });
  process.stdout.write(`${JSON.stringify({ ok: true, output, ...evidence.totals })}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
