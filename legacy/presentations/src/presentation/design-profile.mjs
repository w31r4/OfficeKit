const PROFILE_SCHEMA = "office-kit/pptx-design-profile/v1";
const DEFAULT_MAX_ITEMS = 256;

/**
 * Build a bounded, descriptive view of a Presentation. The result is evidence
 * for an Agent choosing a layout or source-derived asset; it is never an edit
 * plan and contains no XML selectors, package paths, or source bytes.
 */
export function buildPresentationDesignProfile(presentation, {
  sourceRevisionSha256,
  maxItems = DEFAULT_MAX_ITEMS,
  includeComponentCandidates = true,
} = {}) {
  if (!presentation || typeof presentation.inspect !== "function") {
    throw new TypeError("Presentation design profile requires a Presentation instance.");
  }
  if (!Number.isSafeInteger(maxItems) || maxItems < 1 || maxItems > DEFAULT_MAX_ITEMS) {
    throw new RangeError(`Presentation design profile maxItems must be an integer from 1 through ${DEFAULT_MAX_ITEMS}.`);
  }
  const sourceSha = typeof sourceRevisionSha256 === "string" && /^[0-9a-f]{64}$/iu.test(sourceRevisionSha256)
    ? sourceRevisionSha256.toLowerCase()
    : undefined;
  const records = parseNdjson(presentation.inspect({
    kind: "deck,theme,layout,slide,shape,textbox,image,table,chart,connector,groupShape,nativeObject",
    maxChars: Infinity,
  }).ndjson);
  const elements = records.filter((record) => ELEMENT_KINDS.has(record?.kind));
  const slides = records.filter((record) => record?.kind === "slide").sort((a, b) => Number(a.slide || 0) - Number(b.slide || 0));
  const theme = records.find((record) => record?.kind === "theme") || {};
  const componentCandidates = sourceSha && includeComponentCandidates
    ? inspectComponentCandidates(presentation, maxItems)
    : { available: false, reason: sourceSha ? "component candidate inspection disabled" : "source-free presentation has no source-bound candidates" };
  return normalize({
    schema: PROFILE_SCHEMA,
    source: {
      sourceBound: Boolean(sourceSha),
      ...(sourceSha ? { revisionSha256: sourceSha } : {}),
    },
    canvas: {
      width: finiteNumber(presentation.slideSize?.width),
      height: finiteNumber(presentation.slideSize?.height),
      aspectRatio: ratio(presentation.slideSize?.width, presentation.slideSize?.height),
    },
    designLanguage: {
      theme: themeEvidence(theme),
      palette: paletteEvidence(theme, elements),
      typography: typographyEvidence(theme, elements),
      density: densityEvidence(slides, elements),
      rhythm: rhythmEvidence(elements, presentation.slideSize, maxItems),
      vectorAssets: vectorAssetEvidence(presentation, maxItems),
    },
    layoutFamilies: layoutFamilies(records, maxItems),
    slideArchetypes: slideArchetypes(slides, elements),
    reusableComponents: reusableComponents(elements, presentation.slideSize, maxItems),
    componentCandidates,
    nativeOpaque: nativeOpaqueEvidence(elements, maxItems),
  });
}

const ELEMENT_KINDS = new Set(["shape", "textbox", "image", "table", "chart", "connector", "groupShape", "nativeObject"]);

function themeEvidence(theme) {
  return {
    id: theme.id || undefined,
    name: theme.name || undefined,
    colorSchemeName: theme.colorSchemeName || undefined,
    sourceBound: theme.source?.sourceBound === true,
    editable: theme.source?.editable === true,
    xmlSha256: /^[0-9a-f]{64}$/iu.test(String(theme.source?.xmlSha256 || ""))
      ? String(theme.source.xmlSha256).toLowerCase()
      : undefined,
  };
}

function parseNdjson(value) {
  return String(value || "").split("\n").filter(Boolean).map((line) => JSON.parse(line));
}

function inspectComponentCandidates(presentation, maxItems) {
  try {
    const records = parseNdjson(presentation.inspect({ kind: "componentCandidate", maxChars: Infinity }).ndjson);
    const statuses = countValues(records.map((record) => record.status || "unknown"));
    const kinds = countValues(records.map((record) => record.descriptor?.kind || "unknown"));
    const blockedReasons = countValues(records.filter((record) => record.status === "blocked").map((record) => record.blockedReason || "unspecified"));
    return {
      available: true,
      total: records.length,
      statuses: sortedCounts(statuses),
      kinds: sortedCounts(kinds),
      blockedReasons: topCounts(blockedReasons, maxItems),
      inspectOnlyCandidateIds: records
        .filter((record) => record.status === "inspect-only")
        .slice(0, maxItems)
        .map((record) => record.candidateId),
    };
  } catch (error) {
    return { available: false, reason: String(error?.message || "component candidate inspection unavailable") };
  }
}

function paletteEvidence(theme, elements) {
  const direct = new Map();
  const scheme = new Map();
  for (const [key, value] of Object.entries(theme.colors || {})) {
    if (isRgb(value)) direct.set(value.toUpperCase(), (direct.get(value.toUpperCase()) || 0) + 1);
    else if (value) scheme.set(String(key), (scheme.get(String(key)) || 0) + 1);
  }
  for (const element of elements) {
    for (const value of [element.fill, element.line?.fill, element.color]) {
      if (isRgb(value)) direct.set(value.toUpperCase(), (direct.get(value.toUpperCase()) || 0) + 1);
      else if (typeof value === "string" && value && !value.startsWith("#")) scheme.set(value, (scheme.get(value) || 0) + 1);
    }
    for (const style of textStyles(element)) {
      if (isRgb(style.color)) direct.set(style.color.toUpperCase(), (direct.get(style.color.toUpperCase()) || 0) + 1);
      else if (style.color) scheme.set(String(style.color), (scheme.get(String(style.color)) || 0) + 1);
    }
  }
  return { direct: topCounts(direct, 24), scheme: topCounts(scheme, 24) };
}

function typographyEvidence(theme, elements) {
  const fonts = new Map();
  const sizes = [];
  for (const value of [theme.fonts?.major, theme.fonts?.minor, theme.fonts?.majorEastAsia, theme.fonts?.minorEastAsia]) {
    if (value && !String(value).startsWith("+")) fonts.set(String(value), (fonts.get(String(value)) || 0) + 1);
  }
  for (const element of elements) {
    for (const style of textStyles(element)) {
      if (style.fontFamily && !String(style.fontFamily).startsWith("+")) fonts.set(String(style.fontFamily), (fonts.get(String(style.fontFamily)) || 0) + 1);
      const size = Number(style.fontSize);
      if (Number.isFinite(size) && size > 0 && size <= 400) sizes.push(size);
    }
  }
  return {
    fonts: topCounts(fonts, 24),
    fontSizePt: {
      min: sizes.length ? round(Math.min(...sizes), 2) : null,
      median: sizes.length ? round(median(sizes), 2) : null,
      max: sizes.length ? round(Math.max(...sizes), 2) : null,
      samples: sizes.length,
    },
  };
}

function textStyles(element) {
  const styles = [];
  for (const paragraph of Array.isArray(element.paragraphs) ? element.paragraphs : []) {
    if (paragraph.style && typeof paragraph.style === "object") styles.push(paragraph.style);
    for (const run of paragraph.runs || []) if (run.style && typeof run.style === "object") styles.push(run.style);
  }
  if (element.style && typeof element.style === "object") styles.push(element.style);
  return styles;
}

function densityEvidence(slides, elements) {
  const perSlide = (slide, mapper) => mapper(elements.filter((element) => Number(element.slide) === Number(slide.slide)));
  const counts = slides.map((slide) => perSlide(slide, (items) => items.length));
  const chars = slides.map((slide) => perSlide(slide, (items) => items.reduce((sum, element) => sum + String(element.text || "").length, 0)));
  const opaque = slides.map((slide) => perSlide(slide, (items) => items.filter((element) => element.kind === "nativeObject").length));
  return { slides: slides.length, elementsPerSlide: summaryStats(counts), textCharsPerSlide: summaryStats(chars), nativeOpaquePerSlide: summaryStats(opaque) };
}

function vectorAssetEvidence(presentation, maxItems) {
  const images = presentation.slides.items.flatMap((slide) => slide.images?.items || [])
    .filter((image) => typeof image.svgDataUrl === "string" && image.svgDataUrl.length > 0);
  const sourceHashes = new Set();
  let editable = 0;
  for (const image of images) {
    const capability = image.svgEditCapability;
    const textCapability = image.svgTextCapability;
    const sourceHash = capability?.sourceSha256 || textCapability?.sourceSha256;
    if (sourceHash) sourceHashes.add(sourceHash);
    if (capability?.supported === true || textCapability?.supported === true) editable += 1;
  }
  return {
    // assetCount is the number of distinct fallback SVG byte streams, while
    // usageCount records how many picture elements point at one.
    assetCount: sourceHashes.size,
    usageCount: images.length,
    editableUsageCount: editable,
    examples: [...sourceHashes].sort().slice(0, maxItems),
  };
}

function rhythmEvidence(elements, slideSize, maxItems) {
  const width = Number(slideSize?.width);
  const height = Number(slideSize?.height);
  if (!(width > 0) || !(height > 0)) return { normalizedUnits: "slide fraction rounded to 0.001", x: [], y: [], widths: [], heights: [] };
  const positions = elements.filter((element) => Array.isArray(element.bbox) && element.bbox.length === 4).map((element) => {
    const [left, top, elementWidth, elementHeight] = element.bbox.map(Number);
    return {
      left: round(left / width, 3), top: round(top / height, 3),
      width: round(elementWidth / width, 3), height: round(elementHeight / height, 3),
    };
  });
  return {
    normalizedUnits: "slide fraction rounded to 0.001",
    x: topCounts(countValues(positions.map((item) => item.left)), maxItems),
    y: topCounts(countValues(positions.map((item) => item.top)), maxItems),
    widths: topCounts(countValues(positions.map((item) => item.width)), maxItems),
    heights: topCounts(countValues(positions.map((item) => item.height)), maxItems),
  };
}

function layoutFamilies(records, maxItems) {
  return records.filter((record) => record.kind === "layoutTemplate" || record.kind === "layout").slice(0, maxItems).map((layout) => ({
    id: layout.id,
    name: layout.name || undefined,
    type: layout.type || "unknown",
    placeholders: Number(layout.effectivePlaceholders ?? layout.placeholders ?? 0),
    placeholderTypes: [...new Set(layout.placeholderTypes || [])].sort(),
  })).sort((a, b) => String(a.id).localeCompare(String(b.id)));
}

function slideArchetypes(slides, elements) {
  const signatures = new Map();
  const records = slides.map((slide) => {
    const items = elements.filter((element) => Number(element.slide) === Number(slide.slide));
    const counts = countValues(items.map((element) => element.kind));
    const signature = [slide.layoutId || "no-layout", ...Object.entries(counts).sort(([a], [b]) => a.localeCompare(b)).map(([kind, count]) => `${kind}:${count}`), `text:${items.some((element) => String(element.text || "")) ? "yes" : "no"}`].join("|");
    const record = { slide: Number(slide.slide), title: String(slide.title || "").slice(0, 120), layoutId: slide.layoutId || null, signature, elementCounts: sortedCounts(counts), textChars: items.reduce((sum, element) => sum + String(element.text || "").length, 0) };
    const family = signatures.get(signature) || [];
    family.push(record.slide);
    signatures.set(signature, family);
    return record;
  });
  return records.map((record) => ({ ...record, familySize: signatures.get(record.signature).length }));
}

function reusableComponents(elements, slideSize, maxItems) {
  const width = Number(slideSize?.width);
  const height = Number(slideSize?.height);
  if (!(width > 0) || !(height > 0)) return [];
  const bySignature = new Map();
  for (const element of elements) {
    if (!Array.isArray(element.bbox) || element.bbox.length !== 4) continue;
    const [left, top, elementWidth, elementHeight] = element.bbox.map(Number);
    const signature = [element.kind, round(left / width, 3), round(top / height, 3), round(elementWidth / width, 3), round(elementHeight / height, 3), element.name || ""].join("|");
    const entries = bySignature.get(signature) || [];
    entries.push({ slide: Number(element.slide), id: element.id });
    bySignature.set(signature, entries);
  }
  return [...bySignature.entries()]
    .filter(([, entries]) => entries.length >= 2)
    .sort(([, left], [, right]) => right.length - left.length || String(left[0]?.id).localeCompare(String(right[0]?.id)))
    .slice(0, maxItems)
    .map(([signature, occurrences]) => ({ signature, occurrences: occurrences.slice(0, maxItems), count: occurrences.length }));
}

function nativeOpaqueEvidence(elements, maxItems) {
  const opaque = elements.filter((element) => element.kind === "nativeObject");
  return {
    count: opaque.length,
    kinds: sortedCounts(countValues(opaque.map((record) => record.nativeKind || "unknown"))),
    examples: opaque.slice(0, maxItems).map((record) => ({ id: record.id, slide: Number(record.slide), nativeKind: record.nativeKind || "unknown", name: record.name || undefined })),
  };
}

function isRgb(value) { return typeof value === "string" && /^#[0-9a-f]{6}$/iu.test(value); }
function finiteNumber(value) { const number = Number(value); return Number.isFinite(number) ? number : null; }
function ratio(width, height) { const a = Number(width); const b = Number(height); return a > 0 && b > 0 ? round(a / b, 5) : null; }
function countValues(values) { const counts = {}; for (const value of values) counts[value] = (counts[value] || 0) + 1; return counts; }
function sortedCounts(counts) { return Object.fromEntries(Object.entries(counts).sort(([a], [b]) => a.localeCompare(b))); }
function topCounts(counts, limit) { return Object.entries(counts instanceof Map ? Object.fromEntries(counts) : counts).sort(([a, left], [b, right]) => Number(right) - Number(left) || String(a).localeCompare(String(b))).slice(0, limit).map(([value, count]) => ({ value, count })); }
function summaryStats(values) { const numbers = values.map(Number).filter(Number.isFinite).sort((a, b) => a - b); return numbers.length ? { min: numbers[0], median: round(median(numbers), 2), max: numbers.at(-1) } : { min: 0, median: 0, max: 0 }; }
function median(values) { const sorted = [...values].sort((a, b) => a - b); const middle = Math.floor(sorted.length / 2); return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2; }
function round(value, digits = 3) { const factor = 10 ** digits; return Math.round(Number(value) * factor) / factor; }
function normalize(value) { return JSON.parse(JSON.stringify(value, (_key, entry) => entry === undefined ? undefined : entry)); }
