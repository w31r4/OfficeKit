import { resolveColorToken } from "../shared/colors.mjs";

const GRADIENT_KINDS = new Set(["linear", "radial"]);
const MAX_GRADIENT_STOPS = 16;

function gradientColor(value, label) {
  const source = typeof value === "string" ? value : value?.color ?? value?.fill;
  const resolved = String(resolveColorToken(source, source) || "").trim().toLowerCase();
  const short = /^#?([0-9a-f]{3})$/i.exec(resolved)?.[1];
  if (short) return `#${[...short].map((character) => character.repeat(2)).join("")}`;
  const full = /^#?([0-9a-f]{6})$/i.exec(resolved)?.[1];
  if (!full) throw new TypeError(`${label} must be a six-digit RGB color or supported color token.`);
  return `#${full}`;
}
/**
 * Normalize the deliberately small gradient profile shared by the public
 * presentation model and the native PresentationML codec.  Keeping this
 * helper independent of the protobuf layer means authored and imported
 * gradients use exactly the same validation rules.
 */
export function normalizePresentationGradientFill(value, label = "Presentation gradient fill") {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new TypeError(`${label} must be a gradient object.`);
  }
  if (value.type != null && value.type !== "gradient") {
    throw new TypeError(`${label}.type must be gradient.`);
  }
  const kind = String(value.kind || "linear");
  if (!GRADIENT_KINDS.has(kind)) throw new TypeError(`${label}.kind must be linear or radial.`);
  if (!Array.isArray(value.stops) || value.stops.length < 2 || value.stops.length > MAX_GRADIENT_STOPS) {
    throw new RangeError(`${label}.stops must contain 2 through ${MAX_GRADIENT_STOPS} entries.`);
  }
  let previous = -Infinity;
  const stops = value.stops.map((stop, index) => {
    if (!stop || typeof stop !== "object" || Array.isArray(stop)) {
      throw new TypeError(`${label}.stops[${index}] must be an object.`);
    }
    const offset = Number(stop.offset);
    if (!Number.isFinite(offset) || offset < 0 || offset > 1) {
      throw new RangeError(`${label}.stops[${index}].offset must be from 0 through 1.`);
    }
    if (offset < previous) throw new RangeError(`${label}.stops must be ordered by offset.`);
    previous = offset;
    const color = gradientColor(stop.color, `${label}.stops[${index}].color`);
    const normalized = { offset, color };
    if (stop.opacity != null) {
      const opacity = Number(stop.opacity);
      if (!Number.isFinite(opacity) || opacity < 0 || opacity > 1) {
        throw new RangeError(`${label}.stops[${index}].opacity must be from 0 through 1.`);
      }
      normalized.opacity = opacity;
    }
    return normalized;
  });
  if (kind === "radial" && value.angle != null) {
    throw new TypeError(`${label}.angle is only valid for linear gradients.`);
  }
  const normalized = { type: "gradient", kind, stops };
  if (kind === "linear" && value.angle != null) {
    const angle = Number(value.angle);
    if (!Number.isFinite(angle)) throw new TypeError(`${label}.angle must be finite.`);
    normalized.angle = ((angle % 360) + 360) % 360;
  }
  return normalized;
}

export function isPresentationGradientFill(value) {
  return Boolean(value && typeof value === "object" && !Array.isArray(value) && value.type === "gradient");
}

/**
 * Return SVG defs plus a paint reference for the model preview.  The native
 * export remains authoritative; this is intentionally only a visual preview
 * and uses a centered radial profile matching the bounded native contract.
 */
export function presentationGradientFillSvg(value, id, label = "Presentation gradient fill") {
  const fill = normalizePresentationGradientFill(value, label);
  const safeId = String(id || "gradient").replace(/[^A-Za-z0-9_.-]/g, "-");
  const stops = fill.stops.map((stop) =>
    `<stop offset="${stop.offset * 100}%" stop-color="${stop.color}"${stop.opacity == null ? "" : ` stop-opacity="${stop.opacity}"`}/>`
  ).join("");
  const gradient = fill.kind === "radial"
    ? `<radialGradient id="${safeId}" cx="50%" cy="50%" r="50%">${stops}</radialGradient>`
    : `<linearGradient id="${safeId}" x1="0%" y1="0%" x2="100%" y2="0%" gradientTransform="rotate(${fill.angle || 0} .5 .5)">${stops}</linearGradient>`;
  return { defs: `<defs>${gradient}</defs>`, paint: `url(#${safeId})` };
}
