import { isXmlSafeText } from "../shared/xml.mjs";

const ACCESSIBILITY_FIELDS = ["title", "description"];
const MAX_ACCESSIBILITY_TEXT_LENGTH = 1_024;
const importedAccessibilityEditable = new WeakMap();

function normalizeText(value, owner, field) {
  if (typeof value !== "string" || !value.length || value.length > MAX_ACCESSIBILITY_TEXT_LENGTH || !isXmlSafeText(value)) {
    throw new TypeError(`${owner} accessibility.${field} must contain 1 through ${MAX_ACCESSIBILITY_TEXT_LENGTH} XML-safe characters.`);
  }
  return value;
}

export function normalizePresentationAccessibility(value, owner = "Presentation object") {
  if (value == null) return undefined;
  return applyAccessibility(undefined, value, owner, false);
}

export function updatePresentationAccessibility(current, update, owner = "Presentation object") {
  return applyAccessibility(current, update, owner, true);
}

export function initializePresentationAccessibility(target, config, owner = "Presentation object") {
  if (config?._officeKitAccessibilityEditable !== undefined) {
    importedAccessibilityEditable.set(target, Boolean(config._officeKitAccessibilityEditable));
  }
  return normalizePresentationAccessibility(config?.accessibility, owner);
}

export function presentationAccessibilityCapability(target) {
  const sourceBound = importedAccessibilityEditable.has(target);
  const editable = !sourceBound || importedAccessibilityEditable.get(target) === true;
  return { sourceBound, editable, addable: editable };
}

export function setPresentationAccessibilityMetadata(target, current, update, owner = "Presentation object") {
  if (!presentationAccessibilityCapability(target).editable) {
    throw new Error(`${owner} accessibility metadata is source-bound and does not match the editable p:cNvPr profile.`);
  }
  return updatePresentationAccessibility(current, update, owner);
}

function applyAccessibility(current, update, owner, partial) {
  if (!update || typeof update !== "object" || Array.isArray(update)) {
    throw new TypeError(`${owner} accessibility${partial ? " metadata update" : ""} must be an object with title and/or description.`);
  }
  const unsupported = Object.keys(update).filter((field) => !ACCESSIBILITY_FIELDS.includes(field));
  if (unsupported.length) throw new TypeError(`${owner} accessibility${partial ? " metadata" : ""} does not support ${unsupported.join(", ")}.`);
  if (partial && !ACCESSIBILITY_FIELDS.some((field) => Object.hasOwn(update, field))) {
    throw new TypeError(`${owner} accessibility metadata update requires title and/or description.`);
  }
  const result = { ...(current || {}) };
  for (const field of ACCESSIBILITY_FIELDS) {
    if (!Object.hasOwn(update, field)) continue;
    if (update[field] == null) {
      if (partial) delete result[field];
    }
    else result[field] = normalizeText(update[field], owner, field);
  }
  return Object.keys(result).length ? result : undefined;
}
