#!/usr/bin/env node

import { createHash } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

import { FileBlob, PresentationFile } from "../src/index.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const PROFILE_SCHEMA = "office-kit/pptx-design-profile/v1";
const MAX_SOURCE_BYTES = 128 * 1024 * 1024;
const MAX_PROFILE_ITEMS = 256;

export async function buildPptxDesignProfile(inputPath, { id } = {}) {
  const absolute = path.resolve(inputPath);
  const bytes = await readFile(absolute);
  if (bytes.byteLength === 0 || bytes.byteLength > MAX_SOURCE_BYTES) {
    throw new RangeError(`PPTX design profile input must be between 1 byte and ${MAX_SOURCE_BYTES} bytes.`);
  }
  const sourceSha256 = sha256(bytes);
  const zip = await JSZip.loadAsync(bytes);
  const partNames = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const partTexts = new Map();
  for (const name of partNames.filter((candidate) => isTextPart(candidate))) {
    partTexts.set(name, await zip.file(name).async("text"));
  }

  const presentation = await PresentationFile.importPptx(new FileBlob(bytes, { type: PPTX_MIME }));
  const inspection = parseNdjson(presentation.inspect({
    kind: ["deck", "theme", "layout", "slide", "shape", "textbox", "image", "table", "chart", "connector", "groupShape", "nativeObject"],
    maxChars: Infinity,
  }).ndjson);
  const elements = inspection.filter((record) => isElementRecord(record));
  const slides = inspection.filter((record) => record.kind === "slide").sort((a, b) => a.slide - b.slide);
  const layouts = inspection.filter((record) => record.kind === "layoutTemplate" || record.kind === "layout");
  const xml = [...partTexts.values()].join("\n");

  const profile = {
    schema: PROFILE_SCHEMA,
    source: {
      id: String(id || path.basename(absolute, path.extname(absolute))).replace(/[^A-Za-z0-9._-]+/g, "-").replace(/^-+|-+$/g, "") || "presentation",
      fileName: path.basename(absolute),
      bytes: bytes.byteLength,
      sha256: sourceSha256,
    },
    evidence: {
      package: packageEvidence(partNames, partTexts, bytes),
      inspectedRecordCount: inspection.length,
      inspectedElementCount: elements.length,
      structuralPartHashes: await structuralPartHashes(zip, partNames),
    },
    canvas: {
      width: presentation.slideSize.width,
      height: presentation.slideSize.height,
      aspectRatio: round(presentation.slideSize.width / presentation.slideSize.height, 5),
    },
    designLanguage: {
      palette: paletteEvidence(xml),
      typography: typographyEvidence(xml),
      density: densityEvidence(slides, elements),
      rhythm: rhythmEvidence(elements, presentation.slideSize),
    },
    layoutFamilies: layoutFamilies(layouts),
    slideArchetypes: slideArchetypes(slides, elements),
    reusableComponents: reusableComponents(elements, presentation.slideSize),
    nativeOpaque: nativeOpaqueEvidence(elements),
  };

  return normalizeProfile(profile);
}

function parseNdjson(value) {
  return String(value || "").split("\n").filter(Boolean).map((line) => JSON.parse(line));
}

function isElementRecord(record) {
  return ["shape", "textbox", "image", "table", "chart", "connector", "groupShape", "nativeObject"].includes(record?.kind);
}

function packageEvidence(partNames, partTexts, bytes) {
  return {
    partCount: partNames.length,
    slideCount: partNames.filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name)).length,
    layoutCount: partNames.filter((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(name)).length,
    masterCount: partNames.filter((name) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(name)).length,
    themeCount: partNames.filter((name) => /^ppt\/theme\/theme\d+\.xml$/i.test(name)).length,
    mediaCount: partNames.filter((name) => /^ppt\/media\//i.test(name)).length,
    svgCount: partNames.filter((name) => /^ppt\/media\/.*\.svg$/i.test(name)).length,
    chartCount: partNames.filter((name) => /^ppt\/charts\/chart\d+\.xml$/i.test(name)).length,
    diagramCount: partNames.filter((name) => /^ppt\/diagrams\/.*\.xml$/i.test(name)).length,
    embeddingCount: partNames.filter((name) => /^ppt\/embeddings\//i.test(name)).length,
    oleCount: partNames.filter((name) => /^ppt\/embeddings\/.*(?:ole|package)/i.test(name)).length,
    notesCount: partNames.filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(name)).length,
    commentCount: partNames.filter((name) => /^ppt\/comments\//i.test(name)).length,
    textParts: partTexts.size,
    sourceHeaderSha256: sha256(bytes.subarray(0, Math.min(bytes.length, 4096))),
  };
}

async function structuralPartHashes(zip, partNames) {
  const selected = partNames.filter((name) =>
    name === "ppt/presentation.xml"
    || /^ppt\/(?:theme|slideMasters|slideLayouts)\/.*\.xml$/i.test(name),
  );
  const result = {};
  for (const name of selected) result[name] = sha256(await zip.file(name).async("uint8array"));
  return result;
}

export function paletteEvidence(xml) {
  const colors = new Map();
  const schemes = new Map();
  for (const match of String(xml).matchAll(/<(?:[\w.-]+:)?(?:srgbClr|scrgbClr)\b[^>]*\bval="([0-9A-Fa-f]{6})"/g)) {
    const key = `#${match[1].toUpperCase()}`;
    colors.set(key, (colors.get(key) || 0) + 1);
  }
  for (const match of String(xml).matchAll(/<(?:[\w.-]+:)?schemeClr\b[^>]*\bval="([\w-]+)"/g)) {
    const key = String(match[1]);
    schemes.set(key, (schemes.get(key) || 0) + 1);
  }
  return {
    direct: topCounts(colors, 12),
    scheme: topCounts(schemes, 12),
    note: "Counts are source XML occurrences, not a promise that every token is visually dominant.",
  };
}

export function typographyEvidence(xml) {
  const fonts = new Map();
  for (const match of String(xml).matchAll(/<(?:[\w.-]+:)?(?:latin|ea|cs)\b[^>]*\btypeface="([^"]+)"/g)) {
    const value = decodeXml(match[1]).trim();
    if (!value || value.startsWith("+") || value.startsWith("-") || value === "Arial") continue;
    fonts.set(value, (fonts.get(value) || 0) + 1);
  }
  const sizes = [];
  for (const match of String(xml).matchAll(/<(?:[\w.-]+:)?(?:rPr|defRPr|endParaRPr)\b[^>]*\bsz="(\d+)"/g)) {
    const size = Number(match[1]) / 100;
    if (size > 0 && size <= 400) sizes.push(size);
  }
  return {
    fonts: topCounts(fonts, 12),
    fontSizePt: {
      min: sizes.length ? round(Math.min(...sizes), 2) : null,
      median: sizes.length ? round(median(sizes), 2) : null,
      max: sizes.length ? round(Math.max(...sizes), 2) : null,
      samples: sizes.length,
    },
  };
}

function densityEvidence(slides, elements) {
  const elementCounts = slides.map((slide) => elements.filter((element) => element.slide === slide.slide).length);
  const textChars = slides.map((slide) => elements
    .filter((element) => element.slide === slide.slide)
    .reduce((sum, element) => sum + String(element.text || "").length, 0));
  const opaqueCounts = slides.map((slide) => elements.filter((element) => element.slide === slide.slide && element.kind === "nativeObject").length);
  return {
    slides: slides.length,
    elementsPerSlide: summaryStats(elementCounts),
    textCharsPerSlide: summaryStats(textChars),
    nativeOpaquePerSlide: summaryStats(opaqueCounts),
  };
}

function rhythmEvidence(elements, slideSize) {
  const positions = elements.filter((element) => Array.isArray(element.bbox) && element.bbox.length === 4).map((element) => {
    const [left, top, width, height] = element.bbox.map(Number);
    return {
      kind: element.kind,
      left: round(left / slideSize.width, 3),
      top: round(top / slideSize.height, 3),
      width: round(width / slideSize.width, 3),
      height: round(height / slideSize.height, 3),
    };
  });
  const x = topCounts(countValues(positions.map((item) => item.left)), 12);
  const y = topCounts(countValues(positions.map((item) => item.top)), 12);
  const widths = topCounts(countValues(positions.map((item) => item.width)), 12);
  const heights = topCounts(countValues(positions.map((item) => item.height)), 12);
  return { normalizedUnits: "slide fraction rounded to 0.001", x, y, widths, heights };
}

function layoutFamilies(layouts) {
  return layouts.map((layout) => ({
    id: layout.id,
    name: layout.name || undefined,
    type: layout.type || "unknown",
    placeholders: Number(layout.effectivePlaceholders ?? layout.placeholders ?? 0),
    placeholderTypes: [...new Set(layout.placeholderTypes || [])].sort(),
  })).sort((a, b) => String(a.id).localeCompare(String(b.id)));
}

function slideArchetypes(slides, elements) {
  const bySignature = new Map();
  const records = slides.map((slide) => {
    const slideElements = elements.filter((element) => element.slide === slide.slide);
    const counts = countValues(slideElements.map((element) => element.kind));
    const signature = [
      slide.layoutId || "no-layout",
      ...Object.entries(counts).sort(([a], [b]) => a.localeCompare(b)).map(([kind, count]) => `${kind}:${count}`),
      `text:${slideElements.reduce((sum, element) => sum + String(element.text || "").length, 0) > 0 ? "yes" : "no"}`,
    ].join("|");
    const archetype = {
      slide: slide.slide,
      title: String(slide.title || "").slice(0, 120),
      layoutId: slide.layoutId || null,
      signature,
      elementCounts: Object.fromEntries(Object.entries(counts).sort(([a], [b]) => a.localeCompare(b))),
      textChars: slideElements.reduce((sum, element) => sum + String(element.text || "").length, 0),
    };
    const group = bySignature.get(signature) || [];
    group.push(slide.slide);
    bySignature.set(signature, group);
    return archetype;
  });
  return records.map((record) => ({ ...record, familySize: bySignature.get(record.signature).length }));
}

function reusableComponents(elements, slideSize) {
  const bySignature = new Map();
  for (const element of elements) {
    if (!Array.isArray(element.bbox) || element.bbox.length !== 4) continue;
    const [left, top, width, height] = element.bbox.map(Number);
    const normalized = [
      round(left / slideSize.width, 3),
      round(top / slideSize.height, 3),
      round(width / slideSize.width, 3),
      round(height / slideSize.height, 3),
    ];
    const signature = [element.kind, ...normalized, element.name || ""].join("|");
    const entries = bySignature.get(signature) || [];
    entries.push({ slide: element.slide, id: element.id });
    bySignature.set(signature, entries);
  }
  return [...bySignature.entries()]
    .filter(([, entries]) => entries.length >= 2)
    .sort(([, left], [, right]) => right.length - left.length || left[0].id.localeCompare(right[0].id))
    .slice(0, MAX_PROFILE_ITEMS)
    .map(([signature, entries]) => ({ signature, occurrences: entries.slice(0, 12), count: entries.length }));
}

function nativeOpaqueEvidence(elements) {
  const records = elements.filter((element) => element.kind === "nativeObject");
  const kinds = countValues(records.map((record) => record.nativeKind || "unknown"));
  return {
    count: records.length,
    kinds: Object.fromEntries(Object.entries(kinds).sort(([a], [b]) => a.localeCompare(b))),
    examples: records.slice(0, 24).map((record) => ({ id: record.id, slide: record.slide, nativeKind: record.nativeKind || "unknown", name: record.name || undefined })),
  };
}

function normalizeProfile(profile) {
  return JSON.parse(JSON.stringify(profile, (_key, value) => value === undefined ? undefined : value));
}

function summaryStats(values) {
  const numbers = values.map(Number).filter(Number.isFinite).sort((a, b) => a - b);
  if (!numbers.length) return { min: 0, median: 0, max: 0 };
  return { min: numbers[0], median: round(median(numbers), 2), max: numbers.at(-1) };
}

function topCounts(counts, limit) {
  const entries = counts instanceof Map ? [...counts.entries()] : Object.entries(counts);
  return entries.sort(([a, left], [b, right]) => right - left || a.localeCompare(b)).slice(0, limit).map(([value, count]) => ({ value, count }));
}

function countValues(values) {
  const counts = {};
  for (const value of values) counts[value] = (counts[value] || 0) + 1;
  return counts;
}

function median(values) {
  const sorted = [...values].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
}

function round(value, digits = 3) {
  const factor = 10 ** digits;
  return Math.round(Number(value) * factor) / factor;
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function isTextPart(name) {
  return /(?:\.xml|\.rels|\.vml|\.txt|\.svg)$/i.test(name);
}

function decodeXml(value) {
  return String(value).replaceAll("&quot;", '"').replaceAll("&apos;", "'").replaceAll("&lt;", "<").replaceAll("&gt;", ">").replaceAll("&amp;", "&");
}

function parseArgs(argv) {
  const inputs = [];
  let output;
  let force = false;
  for (let index = 0; index < argv.length; index += 1) {
    const flag = argv[index];
    if (flag === "--input") {
      const value = argv[++index];
      if (!value) throw new Error("--input requires [id=]path.");
      const separator = value.indexOf("=");
      const inputId = separator > 0 ? value.slice(0, separator) : undefined;
      const inputPath = separator > 0 ? value.slice(separator + 1) : value;
      inputs.push({ id: inputId, path: inputPath });
    } else if (flag === "--output") {
      output = argv[++index];
      if (!output) throw new Error("--output requires a path.");
    } else if (flag === "--force") {
      force = true;
    } else {
      throw new Error(`Unknown option ${flag}.`);
    }
  }
  if (!inputs.length || !output) throw new Error("Usage: pptx-design-profile.mjs --input [id=]file.pptx [--input [id=]file.pptx ...] --output profile.json [--force]");
  return { inputs, output, force };
}

async function main() {
  const { inputs, output, force } = parseArgs(process.argv.slice(2));
  const profiles = [];
  for (const input of inputs) profiles.push(await buildPptxDesignProfile(input.path, { id: input.id }));
  const result = { schema: PROFILE_SCHEMA, profiles };
  await mkdir(path.dirname(path.resolve(output)), { recursive: true });
  await writeFile(path.resolve(output), `${JSON.stringify(result, null, 2)}\n`, { flag: force ? "w" : "wx" });
  process.stdout.write(`${JSON.stringify({ ok: true, output: path.resolve(output), profiles: profiles.length })}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
