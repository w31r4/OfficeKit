import { resolveColorToken } from "../shared/colors.mjs";
import { xmlEscape } from "../shared/xml.mjs";

const SHAPE_LINE_STYLE_ALIASES = new Map([
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

const SHAPE_LINE_PROPERTIES = new Set(["style", "fill", "color", "width"]);

export function normalizePresentationShapeLine(value = {}, name = "Presentation shape line") {
  const source = value == null ? {} : value;
  if (typeof source !== "object" || Array.isArray(source)) throw new TypeError(`${name} must be an object.`);
  const unsupported = Object.keys(source).filter((key) => !SHAPE_LINE_PROPERTIES.has(key));
  if (unsupported.length) throw new RangeError(`${name} uses unsupported properties: ${unsupported.sort().join(", ")}.`);
  const requestedStyle = String(source.style || "solid");
  const style = SHAPE_LINE_STYLE_ALIASES.get(requestedStyle);
  if (!style) throw new RangeError(`${name} style ${requestedStyle} is unsupported.`);
  const width = Number(source.width ?? 1);
  if (!Number.isFinite(width) || width < 0) throw new RangeError(`${name} width must be a non-negative finite number.`);
  return { ...source, style, width };
}

export function presentationShapeLineColor(value, fallback = "#334155") {
  const raw = value?.fill || value?.color || fallback;
  if (typeof raw === "string") return raw;
  return raw?.color || raw?.fill || fallback;
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
  const line = normalizePresentationShapeLine(value, name);
  const color = presentationShapeLineColor(line);
  const hidden = line.style === "none" || color === "none" || color === "transparent";
  const stroke = hidden ? "none" : resolveColorToken(color, "#334155");
  const dash = hidden ? undefined : SVG_DASH_ARRAYS.get(line.style);
  return `stroke="${xmlEscape(stroke)}" stroke-width="${line.width}"${dash ? ` stroke-dasharray="${dash}"` : ""}`;
}

export function presentationFreeLineSvg(value, position, name = "Presentation free line") {
  const frame = presentationFreeLineFrame(position, name);
  const outline = presentationShapeLineSvgAttributes(value, `${name} outline`);
  return `<line x1="${frame.left}" y1="${frame.top}" x2="${frame.left + frame.width}" y2="${frame.top + frame.height}" fill="none" ${outline}/>`;
}
