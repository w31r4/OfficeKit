const PLAN_SCHEMA = "office-kit/pptx-template-plan/v1";
const MAX_SLIDES = 64;
const DEFAULT_MAX_ITEMS = 64;

const ROLE_HINTS = Object.freeze([
  { tokens: ["title", "cover", "opening", "封面", "标题", "开场"], kinds: ["textbox", "image"], position: "first" },
  { tokens: ["agenda", "overview", "目录", "概览"], kinds: ["textbox", "shape"], position: "early" },
  { tokens: ["chart", "data", "metric", "数据", "指标", "趋势"], kinds: ["chart", "table"], position: "any" },
  { tokens: ["table", "comparison", "compare", "对比", "表格"], kinds: ["table", "textbox"], position: "any" },
  { tokens: ["image", "visual", "photo", "图片", "视觉"], kinds: ["image", "shape"], position: "any" },
  { tokens: ["steps", "process", "roadmap", "流程", "步骤", "路线"], kinds: ["shape", "connector", "textbox"], position: "any" },
  { tokens: ["decision", "action", "next", "决策", "行动", "下一步"], kinds: ["textbox", "shape"], position: "last" },
]);

export function buildTemplateGenerationPlan(presentation, {
  profile,
  slides,
  maxItems = DEFAULT_MAX_ITEMS,
} = {}) {
  if (!presentation || typeof presentation.inspect !== "function") {
    throw new TypeError("Template generation planning requires a Presentation instance.");
  }
  if (!profile || profile.schema !== "office-kit/pptx-design-profile/v1") {
    throw new TypeError("Template generation planning requires a design profile from the same presentation.");
  }
  if (!Number.isSafeInteger(maxItems) || maxItems < 1 || maxItems > DEFAULT_MAX_ITEMS) {
    throw new RangeError(`Template generation plan maxItems must be an integer from 1 through ${DEFAULT_MAX_ITEMS}.`);
  }
  const sourceRevisionSha256 = String(profile.source?.revisionSha256 || "").toLowerCase();
  if (profile.source?.sourceBound !== true || !/^[0-9a-f]{64}$/u.test(sourceRevisionSha256)) {
    const error = new Error("Template generation planning requires a trusted imported PPTX source revision.");
    error.code = "presentation_template_plan_source_required";
    throw error;
  }
  if (!Array.isArray(slides) || slides.length < 1 || slides.length > MAX_SLIDES) {
    throw new RangeError(`Template generation plan requires one through ${MAX_SLIDES} slide requests.`);
  }
  const records = parseNdjson(presentation.inspect({
    kind: "slide,textbox,shape,image,table,chart,connector,groupShape",
    maxChars: Infinity,
  }).ndjson);
  const slideRecords = records.filter((record) => record?.kind === "slide").sort((a, b) => Number(a.slide) - Number(b.slide));
  const elementsBySlide = groupBySlide(records.filter((record) => record?.kind !== "slide"));
  const archetypes = new Map((profile.slideArchetypes || []).map((record) => [Number(record.slide), record]));
  const requests = slides.map((request, index) => normalizeRequest(request, index));
  const usedSignatures = new Map();
  const pages = [];
  const rejected = [];
  for (const request of requests) {
    const selection = selectSourceSlide({ presentation, slideRecords, elementsBySlide, archetypes, request, usedSignatures });
    if (!selection) {
      rejected.push({ index: request.index, role: request.role, reason: "no clone-safe source slide with a bounded text target" });
      continue;
    }
    const { slideRecord, archetype, target, score, alternatives } = selection;
    const sourceSlideOrdinal = Number(slideRecord.slide);
    const sourceSlide = presentation.slides.items[sourceSlideOrdinal - 1];
    const assetCandidates = reusableAssets(profile, sourceSlideOrdinal, request, maxItems);
    const requestedChars = request.title.length + request.body.reduce((sum, value) => sum + value.length, 0);
    const sourceSampleChars = target.sampleText.length;
    const fit = requestedChars <= Math.max(sourceSampleChars, 12) ? "likely" : "review-required";
    const page = {
      index: request.index,
      role: request.role,
      content: { title: request.title, body: request.body },
      sourceSlideOrdinal,
      sourceSlideId: sourceSlide?.id,
      frameTarget: target,
      targetRunText: target.sampleText,
      source: {
        slideOrdinal: sourceSlideOrdinal,
        slideId: sourceSlide?.id,
        layoutId: archetype?.layoutId || slideRecord.layoutId || null,
        archetypeSignature: archetype?.signature || null,
        cloneCapability: sourceSlide?.cloneCapability || null,
      },
      frame: {
        target,
        fit: { status: fit, requestedChars, sourceSampleChars, basis: "source run length heuristic; export and layout review remain required" },
        assetCandidates,
      },
      selection: { score, alternatives },
      review: { required: fit !== "likely" || assetCandidates.some((candidate) => candidate.status === "blocked") },
    };
    pages.push(page);
    const signature = archetype?.signature || `slide:${sourceSlideOrdinal}`;
    usedSignatures.set(signature, (usedSignatures.get(signature) || 0) + 1);
  }
  const warnings = [
    "The plan is source-bound evidence, not mutation authority; re-resolve slide ordinals after each export/reimport.",
    "Text fit is a bounded source-run heuristic. Export, validateLayout, and render/review remain required.",
  ];
  if (profile.nativeOpaque?.count) warnings.push(`${profile.nativeOpaque.count} opaque native objects remain outside semantic authoring.`);
  return deepFreeze({
    schema: PLAN_SCHEMA,
    status: rejected.length ? "blocked" : "ready",
    source: { sourceBound: true, revisionSha256: sourceRevisionSha256, slideCount: slideRecords.length },
    canvas: profile.canvas,
    designLanguage: profile.designLanguage,
    constraints: { maxSlides: MAX_SLIDES, cloneSafeOnly: true, maxItems },
    pages,
    rejected,
    warnings,
  });
}

function normalizeRequest(request, index) {
  if (!request || typeof request !== "object" || Array.isArray(request)) {
    throw new TypeError(`Template generation slide request ${index + 1} must be an object.`);
  }
  const supported = new Set(["role", "title", "body", "sourceSlideOrdinal", "archetypeSignature", "preferredKinds", "assetIntent"]);
  const unsupported = Object.keys(request).filter((key) => !supported.has(key));
  if (unsupported.length) throw new TypeError(`Template generation slide request has unsupported fields: ${unsupported.join(", ")}.`);
  const role = String(request.role || "content").trim();
  const title = String(request.title || "").trim();
  const body = Array.isArray(request.body) ? request.body.map((value) => String(value).trim()).filter(Boolean) : request.body ? [String(request.body).trim()] : [];
  if (!title && !body.length) throw new TypeError(`Template generation slide request ${index + 1} needs title or body content.`);
  const sourceSlideOrdinal = request.sourceSlideOrdinal === undefined ? undefined : Number(request.sourceSlideOrdinal);
  if (sourceSlideOrdinal !== undefined && (!Number.isInteger(sourceSlideOrdinal) || sourceSlideOrdinal < 1)) {
    throw new RangeError(`Template generation slide request ${index + 1} sourceSlideOrdinal must be a positive integer.`);
  }
  const preferredKinds = request.preferredKinds === undefined ? [] : Array.isArray(request.preferredKinds) ? request.preferredKinds.map((value) => String(value).trim()).filter(Boolean) : [String(request.preferredKinds).trim()].filter(Boolean);
  return {
    index,
    role,
    title,
    body,
    sourceSlideOrdinal,
    archetypeSignature: request.archetypeSignature ? String(request.archetypeSignature) : undefined,
    preferredKinds,
    assetIntent: request.assetIntent ? String(request.assetIntent).trim() : "",
  };
}

function selectSourceSlide({ presentation, slideRecords, elementsBySlide, archetypes, request, usedSignatures }) {
  const candidates = slideRecords.map((slideRecord) => {
    const slideOrdinal = Number(slideRecord.slide);
    const slide = presentation.slides.items[slideOrdinal - 1];
    const clone = slide?.cloneCapability;
    const archetype = archetypes.get(slideOrdinal);
    const target = findSlideTextTarget(slide, elementsBySlide.get(slideOrdinal) || []);
    if (!clone?.supported || !target) return null;
    if (request.sourceSlideOrdinal !== undefined && request.sourceSlideOrdinal !== slideOrdinal) return null;
    if (request.archetypeSignature && request.archetypeSignature !== archetype?.signature) return null;
    const score = scoreCandidate({ slideOrdinal, slideRecord, archetype, target, request, usedSignatures });
    return { slideRecord, archetype, target, score, alternatives: [] };
  }).filter(Boolean).sort((left, right) => right.score - left.score || Number(left.slideRecord.slide) - Number(right.slideRecord.slide));
  if (!candidates.length) return null;
  const selected = candidates[0];
  selected.alternatives = candidates.slice(1, 4).map((candidate) => ({ slideOrdinal: Number(candidate.slideRecord.slide), score: candidate.score, archetypeSignature: candidate.archetype?.signature || null }));
  return selected;
}

function scoreCandidate({ slideOrdinal, slideRecord, archetype, target, request, usedSignatures }) {
  const hint = roleHint(request.role);
  const kinds = new Set(Object.keys(archetype?.elementCounts || {}));
  const preferredScore = request.preferredKinds.reduce((score, kind) => score + (kinds.has(kind) ? 20 : 0), 0);
  const roleScore = hint ? hint.kinds.reduce((score, kind) => score + (kinds.has(kind) ? 8 : 0), 0) : 0;
  const targetScore = Math.min(target.sampleText.length, 120) / 10;
  const requestedChars = request.title.length + request.body.reduce((sum, value) => sum + value.length, 0);
  const densityScore = Math.max(0, 12 - Math.abs(Number(archetype?.textChars || 0) - requestedChars) / 20);
  const signature = archetype?.signature || `slide:${slideOrdinal}`;
  const diversityScore = usedSignatures.has(signature) ? -12 * usedSignatures.get(signature) : 16;
  const positionScore = hint?.position === "first" ? Math.max(0, 12 - slideOrdinal) : hint?.position === "early" ? Math.max(0, 8 - slideOrdinal) : hint?.position === "last" ? Math.min(slideRecord.slide || 0, 8) : 0;
  return preferredScore + roleScore + targetScore + densityScore + diversityScore + positionScore;
}

function roleHint(role) {
  const normalized = String(role).toLowerCase();
  return ROLE_HINTS.find((hint) => hint.tokens.some((token) => normalized.includes(token.toLowerCase())));
}

function findTextTarget(elements) {
  const candidates = [];
  for (const element of elements) {
    if (!new Set(["textbox", "shape"]).has(element.kind)) continue;
    for (const [paragraphIndex, paragraph] of (element.paragraphs || []).entries()) {
      for (const [runIndex, run] of (paragraph.runs || []).entries()) {
        const text = String(run.text || "").trim();
        if (!text || /^click to add|^单击此处|^添加(?:副标题|小标题|正文)/iu.test(text)) continue;
        candidates.push({ kind: "shape-run", targetId: element.id, shapeName: element.name || undefined, paragraphIndex, runIndex, sampleText: text, fontSize: Number(run.style?.fontSize) || null });
      }
    }
  }
  return candidates.sort((left, right) => (right.fontSize || 0) - (left.fontSize || 0) || right.sampleText.length - left.sampleText.length || String(left.targetId).localeCompare(String(right.targetId)))[0] || null;
}

function findSlideTextTarget(slide, elements) {
  const shapeTarget = findTextTarget(elements);
  if (shapeTarget) return shapeTarget;
  for (let imageIndex = 0; imageIndex < (slide?.images?.items || []).length; imageIndex += 1) {
    const image = slide.images.items[imageIndex];
    if (image?.svgTextCapability?.supported !== true) continue;
    const nodes = image.getSvgTextNodes();
    const node = nodes.find((candidate) => String(candidate.text || "").trim().length >= 3 &&
      !/^click to add|^单击此处|^添加(?:副标题|小标题|正文)/iu.test(String(candidate.text || "").trim())) || nodes[0];
    if (node) {
      return {
        kind: "svg-text",
        imageIndex,
        nodeId: node.id,
        expectedHash: node.expectedHash,
        sampleText: node.text,
        imageName: image.name || undefined,
        fontSize: null,
      };
    }
  }
  return null;
}

function reusableAssets(profile, slideOrdinal, request, maxItems) {
  const preferred = new Set(request.preferredKinds);
  return (profile.reusableComponents || []).filter((candidate) => candidate.occurrences?.some((occurrence) => Number(occurrence.slide) === slideOrdinal))
    .map((candidate) => ({
      signature: candidate.signature,
      kind: String(candidate.signature || "").split("|")[0] || "unknown",
      count: candidate.count,
      status: preferred.size && !preferred.has(String(candidate.signature || "").split("|")[0]) ? "available" : "preferred",
      intent: request.assetIntent || undefined,
    }))
    .sort((left, right) => (right.status === "preferred" ? 1 : 0) - (left.status === "preferred" ? 1 : 0) || right.count - left.count || left.signature.localeCompare(right.signature))
    .slice(0, maxItems);
}

function groupBySlide(records) {
  const grouped = new Map();
  for (const record of records) {
    const slide = Number(record.slide);
    if (!Number.isInteger(slide) || slide < 1) continue;
    const items = grouped.get(slide) || [];
    items.push(record);
    grouped.set(slide, items);
  }
  return grouped;
}

function parseNdjson(value) {
  return String(value || "").split("\n").filter(Boolean).map((line) => JSON.parse(line));
}

function deepFreeze(value) {
  if (!value || typeof value !== "object" || Object.isFrozen(value)) return value;
  Object.freeze(value);
  for (const child of Object.values(value)) deepFreeze(child);
  return value;
}
