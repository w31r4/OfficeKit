import languageSchema from "./ppj-v1.schema.json" with { type: "json" };
import capabilityRegistry from "./capability-registry.json" with { type: "json" };
import { previewSchemaAt, validatePreviewSupport } from "./preview-capabilities.mjs";
import { previewAssessment, previewDiagnostic, previewPath, mergePreviewDiagnostics } from "./preview-diagnostics.mjs";
import { previewFactualErrors } from "./preview-factual-errors.mjs";

const object = (value) => value !== null && typeof value === "object" && !Array.isArray(value);
const ref = (name) => ({ $ref: `#/$defs/${name}` });

// This only selects schema-owned fields. Validation/default resolution remains
// the compiler's responsibility; unknown descendants always remain unassessed.
function schemaParts(schema, node, value, seen = new Set()) {
  if (!object(node) || seen.has(node)) return [];
  seen.add(node);
  if (node.type === "object" && !object(value) || node.type === "array" && !Array.isArray(value)) return [];
  if (node.type === "string" && typeof value !== "string") return [];
  if ((node.type === "number" || node.type === "integer") && typeof value !== "number") return [];
  if (node.type === "boolean" && typeof value !== "boolean" || node.type === "null" && value !== null) return [];
  if (Object.hasOwn(node, "const") && node.const !== value) return [];
  if (object(value)) for (const discriminator of ["type", "kind"]) {
    const expected = node.properties?.[discriminator]?.const;
    if (expected !== undefined && value[discriminator] !== expected) return [];
  }
  const parts = [node];
  if (node.$ref) parts.push(...schemaParts(schema, previewSchemaAt(schema, node.$ref), value, seen));
  for (const keyword of ["allOf", "oneOf", "anyOf"]) {
    for (const candidate of node[keyword] || []) parts.push(...schemaParts(schema, candidate, value, seen));
  }
  return parts;
}

function childSchema(parts, key, isArray) {
  const children = [];
  for (const part of parts) {
    if (isArray) {
      if (object(part.items)) children.push(part.items);
    } else if (Object.hasOwn(part.properties || {}, key)) children.push(part.properties[key]);
    else if (object(part.additionalProperties)) children.push(part.additionalProperties);
  }
  return children.length ? { anyOf: children } : undefined;
}

/** Assess canonical PPJ without changing input or claiming resolved layout. */
export function assessPpjPreviewInput(program, { schema = languageSchema, registry = capabilityRegistry, assets = new Map() } = {}) {
  const support = validatePreviewSupport(registry, schema);
  const rules = new Map(Object.values(support.fields).map((rule) => [previewSchemaAt(schema, rule.schemaRef), rule]));
  const metadata = new Set(Object.values(support.metadata).map((rule) => previewSchemaAt(schema, rule.schemaRef)));
  const sourceRules = new Map(Object.values(support.sourceBound).map((rule) => [previewSchemaAt(schema, rule.schemaRef), rule]));
  const effectBounds = new Set(["shadow", "glow", "softEdge", "reflection", "innerShadow"].map((name) => schema.$defs[name]).filter(Boolean));
  const active = new Set();

  function emit(context, path, value, rule, extra = {}) {
    context.diagnostics.push(previewDiagnostic({
      pageId: context.pageId, id: context.id, path, value,
      status: rule?.status || "partial", reason: rule?.reason || "preview.field.unassessed",
      action: rule ? `Review this field against the ${rule.owner} output; the local preview does not fully represent it.`
        : "No verified preview mapping exists for this field; inspect the input and add a field regression.",
      ...extra,
    }));
  }

  function visit(value, node, path, context, { nonvisual = false, inheritedRule } = {}) {
    const parts = schemaParts(schema, node, value);
    // Source ownership is determined before shape/discriminator selection. Even
    // an incomplete source descriptor must not expose its payload to traversal.
    function sourceRule(candidate, seen = new Set()) {
      if (!object(candidate) || seen.has(candidate)) return undefined;
      seen.add(candidate);
      if (sourceRules.has(candidate)) return sourceRules.get(candidate);
      if (candidate.$ref) return sourceRule(previewSchemaAt(schema, candidate.$ref), seen);
      for (const branch of [...candidate.allOf || [], ...candidate.oneOf || [], ...candidate.anyOf || []]) {
        const found = sourceRule(branch, seen); if (found) return found;
      }
    }
    const source = sourceRule(node);
    if (source) { emit(context, path, value, source, { sensitive: true }); return; }
    if (ArrayBuffer.isView(value) || value instanceof ArrayBuffer) {
      emit(context, path, value, undefined, { sensitive: true }); return;
    }
    if (value && typeof value === "object" && active.has(value)) {
      emit(context, path, undefined, undefined, { status: "unavailable", reason: "preview.input.cyclic", sensitive: true }); return;
    }
    const known = parts.length > 0;
    // Unknown children of recognized metadata are NOT automatically metadata.
    nonvisual = known && (nonvisual || parts.some((part) => metadata.has(part)));
    const localRule = parts.map((part) => rules.get(part)).find(Boolean);
    const rule = known ? localRule || inheritedRule : undefined;
    if (!nonvisual && parts.some((part) => effectBounds.has(part))) emit(context, path, value, undefined, { reason: "preview.bounds.unassessed", action: "Inspect the effect-expanded extent; the preview does not compute these bounds." });
    if (!value || typeof value !== "object") {
      if (!nonvisual && !(known && identityField(path, value, context))) emit(context, path, value, rule);
      return;
    }
    if (!nonvisual && !known) emit(context, path, value);
    else if (!nonvisual && localRule && !representedContainer(path, value, context)) emit(context, path, value, localRule);
    const entries = Object.entries(value);
    if (!entries.length && !nonvisual && !localRule) emit(context, path, value, rule);
    active.add(value);
    for (const [key, child] of entries) {
      const index = Array.isArray(value) ? Number(key) : key;
      const location = previewPath(path, index), childNode = childSchema(parts, key, Array.isArray(value));
      const childParts = schemaParts(schema, childNode, child);
      const isElement = childParts.some((part) => part === schema.$defs.element);
      if (isElement && !nonvisual) context.children.push(assessElement(child, location, context.pageId, context.inherited));
      else visit(child, childNode, location, context, { nonvisual, inheritedRule: rule });
    }
    active.delete(value);
  }

  function identityField(path, value, context) {
    // Numeric coordinates are the only geometry currently copied verbatim.
    if (context.elementPath && path.startsWith(`${context.elementPath}.frame.`)) {
      const field = path.slice(`${context.elementPath}.frame.`.length);
      if (["x", "y", "width", "height"].includes(field)) return typeof value === "number" && Number.isFinite(value);
      if (field === "rotation") return value === 0;
      if (field === "flipH" || field === "flipV") return value === false;
    }
    if (context.elementPath && path === `${context.elementPath}.type`) return Object.hasOwn(support.types, value);
    if (context.elementPath && path === `${context.elementPath}.hidden`) return value === false;
    if (context.provenImage) {
      if (path === `${context.elementPath}.asset`) return typeof value === "string";
      if (path === `${context.elementPath}.fit`) return value === "contain";
      if (path === `${context.elementPath}.opacity`) return typeof value === "number" && value >= 0 && value <= 1;
    }
    if (path === "$.design.canvas.width" || path === "$.design.canvas.height") return typeof value === "number" && Number.isFinite(value) && value > 0;
    return false;
  }

  function representedContainer(path, value, context) {
    if (!context.provenImage) return false;
    if (path === context.elementPath) return true;
    if (path === `${context.elementPath}.frame`) return object(value);
    return false;
  }

  function inlinePngContain(element) {
    if (element?.type !== "image" || element.fit !== "contain") return false;
    const asset = assets.get(element.asset), bytes = asset?.data;
    if (asset?.mimeType !== "image/png" || !(bytes instanceof Uint8Array) || bytes.length < 24) return false;
    if (![137, 80, 78, 71, 13, 10, 26, 10].every((value, i) => bytes[i] === value)) return false;
    // Decoder success remains a publication requirement. This only recognizes
    // a nonempty PNG header plus the state passed unchanged to SVG <image>.
    if (![73, 72, 68, 82].every((value, i) => bytes[12 + i] === value)) return false;
    if (!bytes.slice(16, 20).some(Boolean) || !bytes.slice(20, 24).some(Boolean)) return false;
    return object(element.frame) && ["x", "y", "width", "height"].every((key) => Number.isFinite(element.frame[key]))
      && element.frame.width > 0 && element.frame.height > 0;
  }

  function assessElement(element, path, pageId, inherited = []) {
    const id = pageId !== undefined && typeof element?.id === "string" && element.id ? element.id : undefined;
    const provenImage = support.fields.image.stateHandler === "inline-png-contain" && inlinePngContain(element);
    const context = { pageId, id, elementPath: path, provenImage, diagnostics: [...inherited], children: [], inherited };
    context.diagnostics.push(...previewFactualErrors(element, { path, pageId, id }, support));
    const type = object(element) && Object.hasOwn(support.types, element.type) ? support.types[element.type] : undefined;
    if (!provenImage) emit(context, previewPath(path, "type"), element?.type, type, type ? {} : { status: "opaque", reason: "preview.element.unknown" });
    if (element?.type === "image" && !assets.get(element.asset)?.data?.byteLength) {
      emit(context, `${path}.asset`, element.asset, undefined, { status: "unavailable", reason: "asset-missing", action: "Restore the image asset and rerun preview from the same input snapshot." });
    }
    if (element?.type === "image" && element.fit === undefined) {
      emit(context, `${path}.fit`, undefined, support.fields.image, { reason: "preview.image.default-fit.unassessed" });
    }
    if (element?.text || element?.type === "text") emit(context, `${path}.text`, element.text, undefined, { reason: "preview.bounds.unassessed", action: "Check text extent and overflow after the actual text layout is available." });
    for (const key of ["rotation", "flipH", "flipV"]) {
      const value = element?.frame?.[key];
      if (value !== undefined && value !== 0 && value !== false) emit(context, `${path}.frame.${key}`, value, undefined, { reason: "preview.bounds.unassessed", action: "Check transformed geometry; the preview frame check is axis-aligned only." });
    }
    for (const key of ["childFrame", "mask", "shadow", "glow", "softEdge", "reflection", "innerShadow"]) {
      if (Object.hasOwn(element || {}, key)) emit(context, `${path}.${key}`, element[key], undefined, { reason: "preview.bounds.unassessed", action: "Review the transformed or effect-expanded bounds; the local frame check does not resolve them." });
    }
    visit(element, type ? { $ref: type.schemaRef } : undefined, path, context);
    return previewAssessment({ path, pageId, id, assessed: true, diagnostics: context.diagnostics,
      children: context.children.map((child) => inherit(child, context.diagnostics)) });
  }

  function inherit(child, diagnostics) {
    return previewAssessment({ ...child, assessed: true, diagnostics: [...child.diagnostics, ...diagnostics],
      children: child.children.map((nested) => inherit(nested, diagnostics)) });
  }

  const global = { diagnostics: [], children: [], inherited: [] };
  if (!object(program)) return previewAssessment();
  // Exclude pages from this first pass; program-level unresolved visual state
  // constrains each page and its elements, retaining the original owning path.
  active.add(program);
  for (const [key, value] of Object.entries(program)) {
    if (key !== "pages") visit(value, schema.properties?.[key], previewPath("$", key), global);
  }
  active.delete(program);
  const globalInherited = mergePreviewDiagnostics(global.diagnostics, ...global.children.map((child) => child.diagnostics));
  const pages = [];
  if (!Array.isArray(program.pages) || program.pages.length === 0) {
    emit(global, "$.pages", program.pages, undefined, { reason: "preview.pages.unassessed" });
  } else for (const [index, page] of program.pages.entries()) {
    const path = `$.pages[${index}]`, pageId = typeof page?.id === "string" && page.id ? page.id : undefined;
    const inherited = globalInherited.map((diagnostic) => Object.freeze({ ...diagnostic, ...(pageId === undefined ? {} : { pageId }) }));
    const context = { pageId, diagnostics: [...inherited], children: [], inherited };
    // Page-owned state affects descendants just as unresolved program state does.
    for (const [key, value] of Object.entries(page || {})) {
      if (key !== "elements") visit(value, schema.$defs.page.properties?.[key], previewPath(path, key), context);
    }
    context.inherited = [...context.diagnostics];
    if (!Array.isArray(page?.elements)) emit(context, `${path}.elements`, page?.elements);
    else page.elements.forEach((element, elementIndex) => context.children.push(assessElement(element, `${path}.elements[${elementIndex}]`, pageId, context.inherited)));
    pages.push(previewAssessment({ path, pageId, assessed: true, diagnostics: context.diagnostics, children: context.children }));
  }
  return previewAssessment({ assessed: true, diagnostics: global.diagnostics, children: [...global.children, ...pages] });
}
