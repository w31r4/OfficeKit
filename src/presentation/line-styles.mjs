import { resolveColorToken } from "../shared/colors.mjs";
import { xmlEscape } from "../shared/xml.mjs";

const LINE_STYLE_ALIASES = new Map([
  ["solid", "solid"],
  ["dashed", "dashed"],
  ["dash", "dashed"],
  ["dotted", "dotted"],
  ["dot", "dotted"],
  ["dash-dot", "dash-dot"],
  ["dashDot", "dash-dot"],
  ["dash-dot-dot", "dash-dot-dot"],
  ["longDashDotDot", "dash-dot-dot"],
  ["none", "none"],
]);

const SVG_DASH_ARRAYS = new Map([
  ["dashed", "8 6"],
  ["dotted", "2 4"],
  ["dash-dot", "8 4 2 4"],
  ["dash-dot-dot", "8 3 2 3 2 3"],
]);

const LINE_END_TYPES = new Set(["triangle", "stealth", "diamond", "oval", "arrow"]);
const LINE_END_SIZES = new Set(["sm", "med", "lg"]);
const LINE_END_PROPERTIES = new Set(["type", "width", "length"]);
const LINE_CAPS = new Set(["flat", "round", "square"]);
const LINE_JOINS = new Set(["round", "bevel", "miter"]);
const LINE_PROPERTIES = new Set([
  "style", "fill", "color", "width",
  "head", "tail", "cap", "join",
  "startArrow", "startArrowWidth", "startArrowLength",
  "endArrow", "endArrowWidth", "endArrowLength",
]);

function own(object, key) {
  return Object.prototype.hasOwnProperty.call(object, key);
}

function mergeCompatibleLineEnds(left, right) {
  if (left == null || right == null) return left === right ? left : undefined;
  for (const key of LINE_END_PROPERTIES) {
    if (left[key] != null && right[key] != null && left[key] !== right[key]) return undefined;
  }
  return { ...right, ...left };
}

export function normalizePresentationLineEnd(value, name = "Presentation line end") {
  if (value == null || value === false || value === "none") return undefined;
  const source = value === true ? { type: "triangle" } : typeof value === "string" ? { type: value } : value;
  if (!source || typeof source !== "object" || Array.isArray(source)) {
    throw new TypeError(`${name} must be a line-end object, token, or false.`);
  }
  const unsupported = Object.keys(source).filter((key) => !LINE_END_PROPERTIES.has(key));
  if (unsupported.length) throw new RangeError(`${name} uses unsupported properties: ${unsupported.sort().join(", ")}.`);
  const type = String(source.type || "none");
  if (type === "none") {
    if (source.width != null || source.length != null) throw new RangeError(`${name} cannot set width or length when type is none.`);
    return undefined;
  }
  if (!LINE_END_TYPES.has(type)) throw new RangeError(`${name}.type must be none, triangle, stealth, diamond, oval, or arrow.`);
  const output = { type };
  for (const key of ["width", "length"]) {
    if (source[key] == null) continue;
    const size = String(source[key]);
    if (!LINE_END_SIZES.has(size)) throw new RangeError(`${name}.${key} must be sm, med, or lg.`);
    output[key] = size;
  }
  return output;
}

function flatLineEnd(source, prefix, name) {
  const typeKey = `${prefix}Arrow`;
  const widthKey = `${prefix}ArrowWidth`;
  const lengthKey = `${prefix}ArrowLength`;
  const present = [typeKey, widthKey, lengthKey].some((key) => own(source, key));
  if (!present) return { present: false, value: undefined };
  if (source[typeKey] == null && (source[widthKey] != null || source[lengthKey] != null)) {
    throw new RangeError(`${name} requires ${typeKey} when arrow width or length is present.`);
  }
  return {
    present: true,
    value: normalizePresentationLineEnd({
      type: source[typeKey] ?? "none",
      ...(source[widthKey] == null ? {} : { width: source[widthKey] }),
      ...(source[lengthKey] == null ? {} : { length: source[lengthKey] }),
    }, name),
  };
}

function lineEnd(source, nestedKey, flatPrefix, name) {
  const nestedPresent = own(source, nestedKey);
  const nested = nestedPresent ? normalizePresentationLineEnd(source[nestedKey], name) : undefined;
  const flat = flatLineEnd(source, flatPrefix, name);
  if (nestedPresent && flat.present) {
    if (nested === undefined || flat.value === undefined) {
      if (nested !== flat.value) throw new RangeError(`${name} has conflicting ${nestedKey} and ${flatPrefix}Arrow values.`);
      return undefined;
    }
    const merged = mergeCompatibleLineEnds(nested, flat.value);
    if (merged === undefined) throw new RangeError(`${name} has conflicting ${nestedKey} and ${flatPrefix}Arrow values.`);
    return merged;
  }
  return nestedPresent ? nested : flat.value;
}

function normalizedChoice(value, choices, name) {
  if (value == null || value === "") return undefined;
  const normalized = String(value);
  if (!choices.has(normalized)) throw new RangeError(`${name} ${normalized} is unsupported.`);
  return normalized;
}

export function normalizePresentationLineCap(value, name = "Presentation line cap") {
  return normalizedChoice(value, LINE_CAPS, name);
}

export function normalizePresentationLineJoin(value, name = "Presentation line join") {
  return normalizedChoice(value, LINE_JOINS, name);
}

export function normalizePresentationLineStyle(value = {}, options = {}) {
  const source = value == null ? {} : value;
  const name = options.name || "Presentation line";
  if (typeof source !== "object" || Array.isArray(source)) throw new TypeError(`${name} must be an object.`);
  const unsupported = Object.keys(source).filter((key) => !LINE_PROPERTIES.has(key));
  if (unsupported.length) throw new RangeError(`${name} uses unsupported properties: ${unsupported.sort().join(", ")}.`);
  const requestedStyle = String(source.style || options.defaultStyle || "solid");
  const style = LINE_STYLE_ALIASES.get(requestedStyle);
  if (!style) throw new RangeError(`${name} style ${requestedStyle} is unsupported.`);
  const width = Number(source.width ?? options.defaultWidth ?? 1);
  if (!Number.isFinite(width) || width < 0) throw new RangeError(`${name} width must be a non-negative finite number.`);
  const head = lineEnd(source, "head", "start", `${name} head`);
  const tail = lineEnd(source, "tail", "end", `${name} tail`);
  const cap = normalizePresentationLineCap(source.cap, `${name} cap`);
  const join = normalizePresentationLineJoin(source.join, `${name} join`);
  return {
    style,
    width,
    ...(own(source, "fill") ? { fill: source.fill } : {}),
    ...(own(source, "color") ? { color: source.color } : {}),
    ...(head ? { head } : {}),
    ...(tail ? { tail } : {}),
    ...(cap ? { cap } : {}),
    ...(join ? { join } : {}),
  };
}

export function presentationLineColor(value, fallback = "#334155") {
  const raw = value?.fill || value?.color || fallback;
  if (typeof raw === "string") return raw;
  return raw?.color || raw?.fill || fallback;
}

function markerShape(type, stroke) {
  if (type === "diamond") return `<path d="M 0 5 L 5 0 L 10 5 L 5 10 z" fill="${xmlEscape(stroke)}"/>`;
  if (type === "oval") return `<ellipse cx="5" cy="5" rx="4" ry="3" fill="${xmlEscape(stroke)}"/>`;
  if (type === "stealth") return `<path d="M 0 0 L 10 5 L 0 10 L 3 5 z" fill="${xmlEscape(stroke)}"/>`;
  if (type === "arrow") return `<path d="M 0 1 L 10 5 L 0 9 L 3 5 z" fill="${xmlEscape(stroke)}"/>`;
  return `<path d="M 0 0 L 10 5 L 0 10 z" fill="${xmlEscape(stroke)}"/>`;
}

export function presentationLineMarkerDefinition(id, end, stroke) {
  if (!end) return "";
  const scale = { sm: 4, med: 6, lg: 8 }[end.width || "med"];
  const length = { sm: 0.8, med: 1, lg: 1.25 }[end.length || "med"];
  return `<marker id="${id}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="${scale * length}" markerHeight="${scale}" orient="auto-start-reverse">${markerShape(end.type, stroke)}</marker>`;
}

function markerBase(value) {
  const sanitized = String(value || "line").replace(/[^A-Za-z0-9_-]/g, "");
  return sanitized || "line";
}

export function presentationLineSvgStyle(value, options = {}) {
  const line = normalizePresentationLineStyle(value, options);
  const color = presentationLineColor(line, options.fallbackColor || "#334155");
  const hidden = line.style === "none" || color === "none" || color === "transparent";
  const stroke = hidden ? "none" : resolveColorToken(color, options.fallbackColor || "#334155");
  const dash = hidden ? undefined : SVG_DASH_ARRAYS.get(line.style);
  const base = markerBase(options.markerBase || options.name);
  const headId = `${base}-head`;
  const tailId = `${base}-tail`;
  const definitions = hidden ? "" : `${presentationLineMarkerDefinition(headId, line.head, stroke)}${presentationLineMarkerDefinition(tailId, line.tail, stroke)}`;
  const markerStart = !hidden && line.head ? ` marker-start="url(#${headId})"` : "";
  const markerEnd = !hidden && line.tail ? ` marker-end="url(#${tailId})"` : "";
  const cap = line.cap ? ` stroke-linecap="${line.cap === "flat" ? "butt" : line.cap}"` : "";
  const join = line.join ? ` stroke-linejoin="${line.join}"` : "";
  return {
    line,
    definitions,
    attributes: `stroke="${xmlEscape(stroke)}" stroke-width="${line.width}"${dash ? ` stroke-dasharray="${dash}"` : ""}${cap}${join}${markerStart}${markerEnd}`,
  };
}

export function presentationFreeLineFrame(value, name = "Presentation free line") {
  const source = value || {};
  const frame = {
    left: Number(source.left),
    top: Number(source.top),
    width: Number(source.width),
    height: Number(source.height),
  };
  if (![frame.left, frame.top, frame.width, frame.height].every(Number.isFinite) || frame.width < 0 || frame.height < 0) {
    throw new RangeError(`${name} requires finite left/top and non-negative width/height.`);
  }
  if (frame.width === 0 && frame.height === 0) throw new RangeError(`${name} requires at least one positive extent.`);
  return frame;
}

export function presentationShapeLineSvgAttributes(value, name = "Presentation shape line") {
  const style = presentationLineSvgStyle(value, { name });
  if (style.line.head || style.line.tail) {
    throw new RangeError(`${name} arrowheads require geometry line.`);
  }
  return style.attributes;
}

export function presentationFreeLineSvg(value, position, name = "Presentation free line", id = name) {
  const frame = presentationFreeLineFrame(position, name);
  const style = presentationLineSvgStyle(value, { name: `${name} outline`, markerBase: id });
  const definitions = style.definitions ? `<defs>${style.definitions}</defs>` : "";
  return `${definitions}<line x1="${frame.left}" y1="${frame.top}" x2="${frame.left + frame.width}" y2="${frame.top + frame.height}" fill="none" ${style.attributes}/>`;
}
