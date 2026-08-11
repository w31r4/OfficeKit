import { isXmlSafeText } from "../shared/xml.mjs";

const ACCESSIBILITY_FIELDS = ["title", "description", "decorative"];
const MAX_ACCESSIBILITY_TEXT_LENGTH = 1_024;
const importedAccessibilityEditable = new WeakMap();

function normalizeText(value, owner, field) {
  if (typeof value !== "string" || !value.length || value.length > MAX_ACCESSIBILITY_TEXT_LENGTH || !isXmlSafeText(value)) {
    throw new TypeError(`${owner} accessibility.${field} must contain 1 through ${MAX_ACCESSIBILITY_TEXT_LENGTH} XML-safe characters.`);
  }
  return value;
}

function applyAccessibility(current, update, owner, partial) {
  if (!update || typeof update !== "object" || Array.isArray(update)) {
    throw new TypeError(`${owner} accessibility${partial ? " metadata update" : ""} must be an object with title, description, and/or decorative.`);
  }
  const unsupported = Object.keys(update).filter((field) => !ACCESSIBILITY_FIELDS.includes(field));
  if (unsupported.length) throw new TypeError(`${owner} accessibility${partial ? " metadata" : ""} does not support ${unsupported.join(", ")}.`);
  if (partial && !ACCESSIBILITY_FIELDS.some((field) => Object.hasOwn(update, field))) {
    throw new TypeError(`${owner} accessibility metadata update requires title, description, and/or decorative.`);
  }
  const result = { ...(current || {}) };
  for (const field of ACCESSIBILITY_FIELDS) {
    if (!Object.hasOwn(update, field)) continue;
    if (update[field] == null) {
      if (partial) delete result[field];
    } else if (field === "decorative") {
      if (typeof update[field] !== "boolean") throw new TypeError(`${owner} accessibility.decorative must be a boolean.`);
      result[field] = update[field];
    } else result[field] = normalizeText(update[field], owner, field);
  }
  if (result.decorative === true && (result.title !== undefined || result.description !== undefined)) {
    throw new TypeError(`${owner} accessibility cannot combine decorative: true with title or description.`);
  }
  return Object.keys(result).length ? result : undefined;
}

export function normalizeSpreadsheetAccessibility(value, owner = "Worksheet drawing") {
  if (value == null) return undefined;
  return applyAccessibility(undefined, value, owner, false);
}

export function initializeSpreadsheetAccessibility(target, config, owner = "Worksheet drawing") {
  if (config?._officeKitAccessibilityEditable !== undefined) {
    importedAccessibilityEditable.set(target, Boolean(config._officeKitAccessibilityEditable));
  }
  return normalizeSpreadsheetAccessibility(config?.accessibility, owner);
}

export function spreadsheetAccessibilityCapability(target) {
  const sourceBound = importedAccessibilityEditable.has(target);
  const editable = !sourceBound || importedAccessibilityEditable.get(target) === true;
  return { sourceBound, editable, addable: editable };
}

export function setSpreadsheetAccessibilityMetadata(target, current, update, owner = "Worksheet drawing") {
  if (!spreadsheetAccessibilityCapability(target).editable) {
    throw new Error(`${owner} accessibility metadata is source-bound and does not match the editable xdr:cNvPr profile.`);
  }
  return applyAccessibility(current, update, owner, true);
}
