#!/usr/bin/env node

import { createHash } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

import { PresentationFile } from "../src/presentation/index.mjs";
import { FileBlob } from "../src/shared/file-blob.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const PROFILE_SCHEMA = "office-kit/pptx-design-profile/v1";
const MAX_SOURCE_BYTES = 128 * 1024 * 1024;
const MAX_PROFILE_ITEMS = 256;
const MAX_SVG_PROFILE_BYTES = 16 * 1024 * 1024;
const MAX_SVG_PROFILE_NODES = 4_096;

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
  const componentCandidates = parseNdjson(presentation.inspect({
    kind: "componentCandidate",
    maxChars: Infinity,
  }).ndjson);
  const elements = inspection.filter((record) => isElementRecord(record));
  const slides = inspection.filter((record) => record.kind === "slide").sort((a, b) => a.slide - b.slide);
  const layouts = inspection.filter((record) => record.kind === "layoutTemplate" || record.kind === "layout");
  const xml = [...partTexts.values()].join("\n");
  const vectorAssets = await svgAssetEvidence(zip, partNames);

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
      vectorAssets,
    },
    layoutFamilies: layoutFamilies(layouts),
    slideArchetypes: slideArchetypes(slides, elements),
    reusableComponents: reusableComponents(elements, presentation.slideSize),
    componentCandidates: componentCandidateEvidence(componentCandidates),
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

async function svgAssetEvidence(zip, partNames) {
  const assets = [];
  for (const name of partNames.filter((candidate) => /^ppt\/media\/.*\.svg$/i.test(candidate))) {
    const bytes = await zip.file(name).async("uint8array");
    assets.push(inspectSvgAsset(name, bytes));
  }
  assets.sort((left, right) => left.part.localeCompare(right.part));
  return {
    assetCount: assets.length,
    supportedCount: assets.filter((asset) => asset.supported).length,
    blockedCount: assets.filter((asset) => !asset.supported).length,
    textNodeCount: assets.reduce((sum, asset) => sum + asset.textNodeCount, 0),
    textChars: assets.reduce((sum, asset) => sum + asset.textChars, 0),
    assets,
  };
}

function inspectSvgAsset(part, bytes) {
  const sourceSha256 = sha256(bytes);
  if (bytes.byteLength === 0 || bytes.byteLength > MAX_SVG_PROFILE_BYTES) {
    return { part, bytes: bytes.byteLength, sourceSha256, supported: false, blockedReason: "SVG exceeds the profile byte budget", textNodeCount: 0, textChars: 0, textSamples: [], fonts: [], fontSizesPt: [], colors: [] };
  }
  let source;
  try {
    source = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    return { part, bytes: bytes.byteLength, sourceSha256, supported: false, blockedReason: "SVG is not valid UTF-8", textNodeCount: 0, textChars: 0, textSamples: [], fonts: [], fontSizesPt: [], colors: [] };
  }
  if (!/^\s*<svg\b[^>]*>/iu.test(source) || !/<\/svg>\s*$/iu.test(source)) {
    return { part, bytes: bytes.byteLength, sourceSha256, supported: false, blockedReason: "SVG root is not bounded", textNodeCount: 0, textChars: 0, textSamples: [], fonts: [], fontSizesPt: [], colors: [] };
  }
  const blockedReason = svgProfileSafety(source);
  if (blockedReason) {
    return { part, bytes: bytes.byteLength, sourceSha256, supported: false, blockedReason, textNodeCount: 0, textChars: 0, textSamples: [], fonts: [], fontSizesPt: [], colors: [] };
  }
  const nodes = [];
  const fonts = new Map();
  const fontSizes = [];
  const colors = new Map();
  const textPattern = /<(?:(?:[A-Za-z_][\w.-]*):)?text\b(?<attributes>[^>]*)>(?<value>[\s\S]*?)<\/(?:(?:[A-Za-z_][\w.-]*):)?text\s*>/giu;
  const tspanPattern = /<(?:(?:[A-Za-z_][\w.-]*):)?tspan\b(?<attributes>[^>]*)>(?<value>[\s\S]*?)<\/(?:(?:[A-Za-z_][\w.-]*):)?tspan\s*>/giu;
  const parentRanges = [];
  const candidates = [];
  for (const match of source.matchAll(textPattern)) {
    const value = match.groups?.value || "";
    const start = match.index ?? -1;
    const openEnd = start + (match[0].indexOf(">") + 1);
    const parentAttributes = svgAttributes(match.groups?.attributes || "");
    const children = [...value.matchAll(tspanPattern)];
    parentRanges.push({ start, end: start + match[0].length });
    if (children.length) {
      for (const child of children) {
        const text = svgTextValue(child.groups?.value || "");
        if (!text) continue;
        candidates.push({
          start: openEnd + (child.index ?? 0),
          tag: "tspan",
          attributes: mergeSvgAttributes(parentAttributes, svgAttributes(child.groups?.attributes || "")),
          text,
        });
      }
      continue;
    }
    const text = svgTextValue(value);
    if (text) candidates.push({ start, tag: "text", attributes: parentAttributes, text });
  }
  // Keep bounded standalone tspans while avoiding duplicates already issued
  // from a surrounding text element.
  for (const match of source.matchAll(tspanPattern)) {
    const start = match.index ?? -1;
    if (parentRanges.some((range) => start > range.start && start < range.end)) continue;
    const text = svgTextValue(match.groups?.value || "");
    if (text) candidates.push({ start, tag: "tspan", attributes: svgAttributes(match.groups?.attributes || ""), text });
  }
  candidates.sort((left, right) => left.start - right.start || left.tag.localeCompare(right.tag));
  for (const candidate of candidates) {
    if (nodes.length >= MAX_SVG_PROFILE_NODES) {
      return { part, bytes: bytes.byteLength, sourceSha256, supported: false, blockedReason: "SVG text node budget exceeded", textNodeCount: 0, textChars: 0, textSamples: [], fonts: [], fontSizesPt: [], colors: [] };
    }
    const { attributes, text } = candidate;
    const fontFamily = attributes["font-family"] || styleValue(attributes.style, "font-family");
    if (fontFamily) fonts.set(fontFamily, (fonts.get(fontFamily) || 0) + 1);
    const fontSize = parseSvgFontSize(attributes["font-size"] || styleValue(attributes.style, "font-size"));
    if (fontSize !== undefined) fontSizes.push(fontSize);
    nodes.push({ index: nodes.length, tag: candidate.tag, text: text.slice(0, 240), textSha256: sha256(text) });
  }
  for (const match of source.matchAll(/\b(?:fill|stroke)\s*=\s*(["'])(#[0-9A-Fa-f]{3,8}|rgb\([^"']+\)|[A-Za-z]+)\1/giu)) {
    const value = match[2].toUpperCase();
    colors.set(value, (colors.get(value) || 0) + 1);
  }
  return {
    part,
    bytes: bytes.byteLength,
    sourceSha256,
    supported: true,
    textNodeCount: nodes.length,
    textChars: nodes.reduce((sum, node) => sum + node.text.length, 0),
    textSamples: nodes.slice(0, 12),
    fonts: topCounts(fonts, 12),
    fontSizesPt: topCounts(countValues(fontSizes.map((value) => String(value))), 12).map(({ value, count }) => ({ value: Number(value), count })),
    colors: topCounts(colors, 12),
  };
}

function svgProfileSafety(source) {
  if (/<\s*(?:script|foreignObject|iframe|object|embed)\b/iu.test(source) ||
      /<!\s*(?:DOCTYPE|ENTITY)\b/iu.test(source) ||
      /\bon[A-Za-z][\w.-]*\s*=/u.test(source) ||
      /(?:href|xlink:href)\s*=\s*(["'])(?!#|data:image\/)[^"']+\1/iu.test(source)) {
    return "SVG contains active content or an external reference";
  }
  return "";
}

function svgAttributes(value) {
  return Object.fromEntries([...String(value).matchAll(/([A-Za-z_:][\w:.-]*)\s*=\s*(["'])(.*?)\2/gu)].map((match) => [match[1].toLowerCase(), decodeXml(match[3])]));
}

function mergeSvgAttributes(parent, child) {
  const merged = { ...parent, ...child };
  if (parent.style && child.style) merged.style = `${parent.style};${child.style}`;
  return merged;
}

function svgTextValue(value) {
  return decodeXml(String(value).replace(/<[^>]*>/gu, "")).replace(/\s+/gu, " ").trim();
}

function styleValue(style, key) {
  const match = String(style || "").match(new RegExp(`(?:^|;)\\s*${key}\\s*:\\s*([^;]+)`, "iu"));
  return match?.[1]?.trim() || "";
}

function parseSvgFontSize(value) {
  const match = String(value || "").trim().match(/^([0-9]+(?:\.[0-9]+)?)\s*(pt|px)?$/iu);
  if (!match) return undefined;
  const number = Number(match[1]);
  if (!Number.isFinite(number) || number <= 0 || number > 400) return undefined;
  return round(match[2]?.toLowerCase() === "px" ? number * 0.75 : number, 2);
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

function componentCandidateEvidence(records) {
  const statuses = countValues(records.map((record) => record.status || "unknown"));
  const kinds = countValues(records.map((record) => record.descriptor?.kind || "unknown"));
  const blockedReasons = countValues(records
    .filter((record) => record.status === "blocked")
    .map((record) => record.blockedReason || "unspecified"));
  const occurrenceCounts = records.map((record) => record.occurrences?.length || 0);
  return {
    total: records.length,
    statuses: Object.fromEntries(Object.entries(statuses).sort(([a], [b]) => a.localeCompare(b))),
    kinds: Object.fromEntries(Object.entries(kinds).sort(([a], [b]) => a.localeCompare(b))),
    occurrenceCount: summaryStats(occurrenceCounts),
    blockedReasons: topCounts(blockedReasons, 12),
    inspectOnlyCandidateIds: records.filter((record) => record.status === "inspect-only").slice(0, 24).map((record) => record.candidateId),
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
