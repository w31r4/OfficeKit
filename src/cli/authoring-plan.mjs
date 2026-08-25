import { createHash } from "node:crypto";

export const PRESENTATION_AUTHORING_PLAN_SCHEMA = "office-kit/presentation-authoring-plan/v1";
export const MAX_AUTHORING_PLAN_BYTES = 256 * 1024;
export const MAX_AUTHORING_PLAN_PAGES = 64;

export const PRESENTATION_AUTHORING_MODES = Object.freeze([
  "create",
  "create-from-template",
  "edit-existing",
  "continue",
]);

export const PRESENTATION_DESIGN_SOURCE_MODES = Object.freeze([
  "self-directed",
  "design-system",
  "template",
  "style-transfer",
]);

export const PRESENTATION_DESIGN_MECHANISMS = Object.freeze([
  "editorial-minimal",
  "enterprise-data-review",
  "technical-architecture",
  "visual-narrative",
  "academic-research",
  "brand-launch",
]);

export const PRESENTATION_DELIVERY_MODES = Object.freeze(["live", "reader", "hybrid"]);
export const PRESENTATION_MOTION_POLICIES = Object.freeze(["adaptive", "none", "explicit"]);
export const PRESENTATION_MOTION_RECIPES = Object.freeze([
  "data-rise",
  "causal-reveal",
  "comparison-beat",
  "focus-pulse",
  "calm-continuity",
  "morph-continuity",
]);

const MODE_SET = new Set(PRESENTATION_AUTHORING_MODES);
const SOURCE_MODE_SET = new Set(PRESENTATION_DESIGN_SOURCE_MODES);
const MECHANISM_SET = new Set(PRESENTATION_DESIGN_MECHANISMS);
const DELIVERY_MODE_SET = new Set(PRESENTATION_DELIVERY_MODES);
const MOTION_POLICY_SET = new Set(PRESENTATION_MOTION_POLICIES);
const MOTION_RECIPE_SET = new Set(PRESENTATION_MOTION_RECIPES);
const MOTION_PURPOSE_RECIPE = new Map([
  ["data-reveal", "data-rise"],
  ["causal-sequence", "causal-reveal"],
  ["comparison", "comparison-beat"],
  ["focus", "focus-pulse"],
  ["continuity", "calm-continuity"],
  ["morph", "morph-continuity"],
]);
const MOTION_TRANSITIONS = new Set(["none", "fade", "push", "morph"]);
const MOTION_STARTS = new Set(["withPrevious", "afterPrevious", "onClick"]);
const TOP_LEVEL_KEYS = new Set([
  "schema",
  "mode",
  "brief",
  "narrative",
  "design",
  "pages",
  "editorial",
  "artifactRefs",
  "recipe",
  "unresolved",
  "nextAction",
]);
const FORBIDDEN_KEY = /^(?:raw[-_]?xml|xml|xpath|part[-_]?path|relationship[-_]?id|relationshipid|r:id|ooxml|selector)$/iu;
const SAFE_ID = /^[a-z0-9][a-z0-9._-]{0,127}$/u;
const SHA256 = /^[a-f0-9]{64}$/u;

export function normalizePresentationAuthoringPlan(value) {
  assertPlainJson(value, "$", new Set(), 0);
  assertExactKeys(value, TOP_LEVEL_KEYS, "Authoring plan");
  if (value.schema !== PRESENTATION_AUTHORING_PLAN_SCHEMA) {
    throw planError("unsupported-plan-schema", `Authoring plan schema must be ${PRESENTATION_AUTHORING_PLAN_SCHEMA}.`);
  }
  if (!MODE_SET.has(value.mode)) {
    throw planError("invalid-authoring-plan", `Authoring plan mode must be one of ${PRESENTATION_AUTHORING_MODES.join(", ")}.`);
  }
  for (const field of ["brief", "narrative", "design", "editorial"]) {
    if (!isPlainObject(value[field])) throw planError("invalid-authoring-plan", `Authoring plan ${field} must be a plain object.`);
  }
  validateBrief(value.brief);
  validateDesign(value.design);
  validatePages(value.pages, value.design.motionPolicy);
  const artifactRefs = validateArtifactRefs(value.artifactRefs ?? []);
  validateReferences(value, artifactRefs);
  boundedString(value.recipe, "Authoring plan recipe", 160);
  if (!Array.isArray(value.unresolved) || value.unresolved.length > 128) {
    throw planError("invalid-authoring-plan", "Authoring plan unresolved must be an array of at most 128 items.");
  }
  if (value.nextAction != null) boundedString(value.nextAction, "Authoring plan nextAction", 1_024);

  const canonical = canonicalJson(value);
  const bytes = Buffer.from(canonical, "utf8");
  if (bytes.byteLength > MAX_AUTHORING_PLAN_BYTES) {
    throw planError("authoring-plan-too-large", `Authoring plan exceeds ${MAX_AUTHORING_PLAN_BYTES} bytes.`);
  }
  return Object.freeze({
    plan: structuredClone(value),
    canonical,
    bytes,
    sha256: createHash("sha256").update(bytes).digest("hex"),
    pageCount: value.pages.length,
    deliveryMode: value.brief.deliveryMode ?? "hybrid",
    motionPolicy: value.design.motionPolicy ?? "adaptive",
    motionPageCount: value.pages.filter((page) => page.motionIntent != null).length,
    designGrammarSha256: createHash("sha256").update(canonicalJson(value.design.designGrammar)).digest("hex"),
  });
}

export function authoringPlanDescriptor(normalized, { path = null, state = "working" } = {}) {
  if (!normalized || !SHA256.test(normalized.sha256) || !Number.isSafeInteger(normalized.pageCount)) {
    throw planError("invalid-authoring-plan", "Cannot describe an invalid authoring plan.");
  }
  return Object.freeze({
    schema: normalized.plan.schema,
    mode: normalized.plan.mode,
    pageCount: normalized.pageCount,
    recipe: normalized.plan.recipe,
    deliveryMode: normalized.deliveryMode,
    motionPolicy: normalized.motionPolicy,
    motionPageCount: normalized.motionPageCount,
    designGrammarSha256: normalized.designGrammarSha256,
    state,
    sha256: normalized.sha256,
    bytes: normalized.bytes.byteLength,
    path,
  });
}

export function canonicalJson(value) {
  return `${canonicalValue(value)}\n`;
}

function validateBrief(brief) {
  if (brief.deliveryMode != null && !DELIVERY_MODE_SET.has(brief.deliveryMode)) {
    throw planError("invalid-authoring-plan", `Authoring plan brief.deliveryMode must be one of ${PRESENTATION_DELIVERY_MODES.join(", ")}.`);
  }
}

function validateDesign(design) {
  if (!SOURCE_MODE_SET.has(design.sourceMode)) {
    throw planError("invalid-authoring-plan", `Authoring plan design.sourceMode must be one of ${PRESENTATION_DESIGN_SOURCE_MODES.join(", ")}.`);
  }
  if (!Array.isArray(design.mechanismPacks) || design.mechanismPacks.length > 2) {
    throw planError("invalid-authoring-plan", "Authoring plan design.mechanismPacks must contain zero to two entries.");
  }
  if (new Set(design.mechanismPacks).size !== design.mechanismPacks.length || design.mechanismPacks.some((entry) => !MECHANISM_SET.has(entry))) {
    throw planError("invalid-authoring-plan", `Authoring plan design mechanisms must be unique members of ${PRESENTATION_DESIGN_MECHANISMS.join(", ")}.`);
  }
  if (!isPlainObject(design.designGrammar) || Object.keys(design.designGrammar).length === 0) {
    throw planError("invalid-authoring-plan", "Authoring plan design.designGrammar must be a non-empty plain object.");
  }
  if (design.motionPolicy != null && !MOTION_POLICY_SET.has(design.motionPolicy)) {
    throw planError("invalid-authoring-plan", `Authoring plan design.motionPolicy must be one of ${PRESENTATION_MOTION_POLICIES.join(", ")}.`);
  }
}

function validatePages(pages, motionPolicy) {
  if (!Array.isArray(pages) || pages.length === 0 || pages.length > MAX_AUTHORING_PLAN_PAGES) {
    throw planError("invalid-authoring-plan", `Authoring plan pages must contain 1 to ${MAX_AUTHORING_PLAN_PAGES} entries.`);
  }
  const ids = new Set();
  for (const [index, page] of pages.entries()) {
    if (!isPlainObject(page)) throw planError("invalid-authoring-plan", `Authoring plan page ${index + 1} must be a plain object.`);
    const id = boundedString(page.id, `Authoring plan page ${index + 1} id`, 128);
    if (!SAFE_ID.test(id) || ids.has(id)) throw planError("invalid-authoring-plan", "Authoring plan page IDs must be unique safe identifiers.");
    ids.add(id);
    boundedString(page.readerTask, `Authoring plan page ${id} readerTask`, 1_024);
    boundedString(page.claim, `Authoring plan page ${id} claim`, 4_096);
    boundedString(page.compositionIntent, `Authoring plan page ${id} compositionIntent`, 2_048);
    if (page.contentBudget != null) validateContentBudget(page.contentBudget, id);
    if (page.motionIntent != null) validateMotionIntent(page.motionIntent, id, motionPolicy);
  }
}

function validateMotionIntent(value, pageId, motionPolicy) {
  if (!isPlainObject(value)) throw planError("invalid-authoring-plan", `Authoring plan page ${pageId} motionIntent must be a plain object.`);
  assertExactKeys(value, new Set(["purpose", "recipe", "units", "transition"]), `Authoring plan page ${pageId} motionIntent`);
  if (!MOTION_PURPOSE_RECIPE.has(value.purpose)) {
    throw planError("invalid-authoring-plan", `Authoring plan page ${pageId} motionIntent purpose is invalid.`);
  }
  if (!MOTION_RECIPE_SET.has(value.recipe) || MOTION_PURPOSE_RECIPE.get(value.purpose) !== value.recipe) {
    throw planError("invalid-authoring-plan", `Authoring plan page ${pageId} motionIntent recipe does not match its purpose.`);
  }
  if (!MOTION_TRANSITIONS.has(value.transition ?? "none")) {
    throw planError("invalid-authoring-plan", `Authoring plan page ${pageId} motionIntent transition is invalid.`);
  }
  if (!Array.isArray(value.units) || value.units.length > 32) {
    throw planError("invalid-authoring-plan", `Authoring plan page ${pageId} motionIntent units must contain at most 32 entries.`);
  }
  if (motionPolicy === "none" && (value.units.length > 0 || value.transition !== "none")) {
    throw planError("invalid-authoring-plan", `Authoring plan page ${pageId} cannot declare motion while design.motionPolicy is none.`);
  }
  const ids = new Set();
  const orders = new Set();
  for (const [index, unit] of value.units.entries()) {
    if (!isPlainObject(unit)) throw planError("invalid-authoring-plan", `Authoring plan page ${pageId} motion unit ${index + 1} must be a plain object.`);
    assertExactKeys(unit, new Set(["id", "targetRole", "order", "start"]), `Authoring plan page ${pageId} motion unit ${index + 1}`);
    const id = boundedString(unit.id, `Authoring plan page ${pageId} motion unit ${index + 1} id`, 128);
    boundedString(unit.targetRole, `Authoring plan page ${pageId} motion unit ${id} targetRole`, 256);
    if (!SAFE_ID.test(id) || ids.has(id)) throw planError("invalid-authoring-plan", `Authoring plan page ${pageId} motion unit IDs must be unique safe identifiers.`);
    if (!Number.isSafeInteger(unit.order) || unit.order <= 0 || orders.has(unit.order)) {
      throw planError("invalid-authoring-plan", `Authoring plan page ${pageId} motion unit order must be unique positive integers.`);
    }
    if (unit.start != null && !MOTION_STARTS.has(unit.start)) {
      throw planError("invalid-authoring-plan", `Authoring plan page ${pageId} motion unit start is invalid.`);
    }
    ids.add(id);
    orders.add(unit.order);
  }
}

function validateContentBudget(value, pageId) {
  if (!isPlainObject(value)) throw planError("invalid-authoring-plan", `Authoring plan page ${pageId} contentBudget must be a plain object.`);
  for (const [key, maximum] of [["maxCharacters", 100_000], ["maxWords", 20_000], ["maxObjects", 10_000]]) {
    if (value[key] != null && (!Number.isSafeInteger(value[key]) || value[key] <= 0 || value[key] > maximum)) {
      throw planError("invalid-authoring-plan", `Authoring plan page ${pageId} contentBudget.${key} is invalid.`);
    }
  }
}

function validateArtifactRefs(value) {
  if (!Array.isArray(value) || value.length > 128) {
    throw planError("invalid-authoring-plan", "Authoring plan artifactRefs must be an array of at most 128 entries.");
  }
  const refs = new Map();
  for (const [index, ref] of value.entries()) {
    if (!isPlainObject(ref)) throw planError("invalid-authoring-plan", `Authoring plan artifactRef ${index + 1} must be a plain object.`);
    const artifactId = boundedString(ref.artifactId, `Authoring plan artifactRef ${index + 1} artifactId`, 64);
    if (!SAFE_ID.test(artifactId) || !SHA256.test(ref.sha256)) {
      throw planError("invalid-authoring-plan", "Authoring plan artifact references require a safe artifactId and lowercase SHA-256.");
    }
    const key = `${artifactId}:${ref.sha256}`;
    if (refs.has(key)) throw planError("invalid-authoring-plan", "Authoring plan artifact references must be unique.");
    if (ref.role != null) boundedString(ref.role, `Authoring plan artifactRef ${artifactId} role`, 128);
    if (ref.label != null) boundedString(ref.label, `Authoring plan artifactRef ${artifactId} label`, 255);
    refs.set(key, ref);
  }
  return refs;
}

function validateReferences(value, declaredRefs) {
  walkJson(value, (entry, key, location) => {
    if (FORBIDDEN_KEY.test(key)) {
      throw planError("unsafe-authoring-plan", `Authoring plan field ${location} exposes raw OOXML state.`);
    }
    if (key === "artifactRef") {
      if (!isPlainObject(entry) || !SAFE_ID.test(entry.artifactId) || !SHA256.test(entry.sha256)) {
        throw planError("invalid-authoring-plan", `Authoring plan field ${location} must contain artifactId and SHA-256.`);
      }
      if (!declaredRefs.has(`${entry.artifactId}:${entry.sha256}`)) {
        throw planError("invalid-authoring-plan", `Authoring plan field ${location} is not declared in artifactRefs.`);
      }
    }
  });
}

function walkJson(value, visit, location = "$") {
  if (Array.isArray(value)) {
    value.forEach((entry, index) => walkJson(entry, visit, `${location}[${index}]`));
    return;
  }
  if (!isPlainObject(value)) return;
  for (const [key, entry] of Object.entries(value)) {
    visit(entry, key, `${location}.${key}`);
    walkJson(entry, visit, `${location}.${key}`);
  }
}

function assertPlainJson(value, location, ancestors, depth) {
  if (depth > 64) throw planError("invalid-authoring-plan", `Authoring plan nesting exceeds 64 levels at ${location}.`);
  if (value === null || typeof value === "string" || typeof value === "boolean") return;
  if (typeof value === "number") {
    if (!Number.isFinite(value)) throw planError("invalid-authoring-plan", `Authoring plan contains a non-finite number at ${location}.`);
    return;
  }
  if (typeof value !== "object") throw planError("invalid-authoring-plan", `Authoring plan contains a non-JSON value at ${location}.`);
  if (ancestors.has(value)) throw planError("invalid-authoring-plan", `Authoring plan contains a cycle at ${location}.`);
  if (!Array.isArray(value) && !isPlainObject(value)) throw planError("invalid-authoring-plan", `Authoring plan contains a non-plain object at ${location}.`);
  ancestors.add(value);
  if (Array.isArray(value)) {
    if (value.length > 10_000) throw planError("invalid-authoring-plan", `Authoring plan array is too large at ${location}.`);
    value.forEach((entry, index) => assertPlainJson(entry, `${location}[${index}]`, ancestors, depth + 1));
  } else {
    if (Object.keys(value).length > 1_000) throw planError("invalid-authoring-plan", `Authoring plan object is too large at ${location}.`);
    for (const [key, entry] of Object.entries(value)) {
      if (key.length === 0 || key.length > 256 || /[\u0000-\u001f\u007f]/u.test(key)) {
        throw planError("invalid-authoring-plan", `Authoring plan contains an invalid key at ${location}.`);
      }
      assertPlainJson(entry, `${location}.${key}`, ancestors, depth + 1);
    }
  }
  ancestors.delete(value);
}

function canonicalValue(value) {
  if (value === null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(canonicalValue).join(",")}]`;
  return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${canonicalValue(value[key])}`).join(",")}}`;
}

function assertExactKeys(value, allowed, label) {
  const unknown = Object.keys(value).filter((key) => !allowed.has(key));
  if (unknown.length > 0) throw planError("invalid-authoring-plan", `${label} contains undeclared field ${unknown[0]}.`);
}

function boundedString(value, label, maximum) {
  if (typeof value !== "string" || value.trim() === "" || value.length > maximum || /[\u0000-\u001f\u007f]/u.test(value)) {
    throw planError("invalid-authoring-plan", `${label} must be non-empty bounded text.`);
  }
  return value;
}

function isPlainObject(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

export function planError(code, message) {
  const error = new Error(message);
  error.code = code;
  return error;
}
