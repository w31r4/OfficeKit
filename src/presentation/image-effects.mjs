const BORDER_STYLES = new Map([
  ["solid", "solid"],
  ["dashed", "dashed"],
  ["dash", "dashed"],
  ["dotted", "dotted"],
  ["dot", "dotted"],
  ["dash-dot", "dash-dot"],
  ["dashDot", "dash-dot"],
  ["dash-dot-dot", "dash-dot-dot"],
  ["longDashDotDot", "dash-dot-dot"],
]);
const BORDER_CAPS = new Set(["flat", "round", "square"]);
const BORDER_JOINS = new Set(["miter", "round", "bevel"]);
const BORDER_KEYS = new Set(["color", "fill", "width", "style", "dash", "cap", "join", "opacity"]);
const SHADOW_KEYS = new Set(["color", "fill", "colorScheme", "scheme", "blurRadius", "blur", "distance", "direction", "angle", "opacity", "alignment", "rotateWithShape"]);
const SHADOW_SCHEME_COLORS = new Set([
  "dk1", "lt1", "dk2", "lt2", "tx1", "bg1", "tx2", "bg2",
  "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink",
]);
const SHADOW_ALIGNMENTS = new Set(["tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"]);
const SHADOW_PRESETS = Object.freeze({
  "shadow-sm": { color: "#000000", blurRadius: 4, distance: 2, direction: 45, opacity: 0.15 },
  shadow: { color: "#000000", blurRadius: 6, distance: 3, direction: 45, opacity: 0.18 },
  "shadow-md": { color: "#000000", blurRadius: 10, distance: 4, direction: 45, opacity: 0.2 },
  "shadow-lg": { color: "#000000", blurRadius: 15, distance: 6, direction: 45, opacity: 0.22 },
  "shadow-xl": { color: "#000000", blurRadius: 22, distance: 9, direction: 45, opacity: 0.24 },
  "shadow-2xl": { color: "#000000", blurRadius: 32, distance: 14, direction: 45, opacity: 0.25 },
});

// DrawingML uses a finite vocabulary for preset geometry.  Keeping the
// vocabulary here (rather than accepting arbitrary XML names) lets an
// imported picture expose a useful mask without turning the public model into
// a raw geometry escape hatch.  The list mirrors the bounded native picture
// reader; "rect" is represented by undefined in the JS model because it is
// the native default.
const IMAGE_MASK_PRESETS = new Set([
  "rect", "ellipse", "roundRect", "line", "triangle", "rightTriangle", "diamond",
  "parallelogram", "trapezoid", "pentagon", "hexagon", "heptagon", "octagon", "chevron",
  "homePlate", "pie", "arc", "donut", "blockArc", "heart", "lightningBolt", "sun", "moon",
  "cloud", "star4", "star5", "star6", "star8", "star10", "star12", "leftArrow", "rightArrow",
  "upArrow", "downArrow", "leftRightArrow", "upDownArrow", "quadArrow", "bentArrow",
  "uturnArrow", "circularArrow", "wedgeRoundRectCallout", "wedgeEllipseCallout", "bracePair",
  "bracketPair", "flowChartProcess", "flowChartDecision", "flowChartData", "flowChartTerminator",
  "flowChartDocument", "flowChartPreparation",
]);

function assertObject(value, label) {
  if (value == null || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${label} must be an object.`);
  return value;
}

function boundedNumber(value, label, { min = 0, max = Infinity } = {}) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < min || number > max) throw new RangeError(`${label} must be from ${min} through ${max}.`);
  return number;
}

function normalizeColor(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label}.color must be a non-empty color string.`);
  const raw = value.trim();
  const short = /^#?([0-9a-f]{3})$/i.exec(raw)?.[1];
  if (short) return `#${[...short].map((char) => char.repeat(2)).join("")}`;
  return raw;
}

export function normalizePresentationImageBorder(value, label = "Presentation image border") {
  if (value == null || value === false || value === "none") return undefined;
  const source = assertObject(value, label);
  const unsupported = Object.keys(source).filter((key) => !BORDER_KEYS.has(key));
  if (unsupported.length) throw new RangeError(`${label} uses unsupported properties: ${unsupported.sort().join(", ")}.`);
  const color = source.color ?? source.fill ?? "#000000";
  const normalizedColor = normalizeColor(color, label);
  const requestedStyle = source.style ?? source.dash ?? "solid";
  const style = BORDER_STYLES.get(String(requestedStyle));
  if (!style) throw new RangeError(`${label}.style must be solid, dashed, dotted, dash-dot, or dash-dot-dot.`);
  const width = boundedNumber(source.width ?? 1, `${label}.width`);
  const cap = source.cap == null || source.cap === "" ? undefined : String(source.cap);
  if (cap !== undefined && !BORDER_CAPS.has(cap)) throw new RangeError(`${label}.cap must be flat, round, or square.`);
  const join = source.join == null || source.join === "" ? undefined : String(source.join);
  if (join !== undefined && !BORDER_JOINS.has(join)) throw new RangeError(`${label}.join must be miter, round, or bevel.`);
  const opacity = source.opacity == null ? undefined : boundedNumber(source.opacity, `${label}.opacity`, { max: 1 });
  return {
    color: normalizedColor,
    width,
    style,
    ...(cap ? { cap } : {}),
    ...(join ? { join } : {}),
    ...(opacity === undefined ? {} : { opacity }),
  };
}

export function normalizePresentationImageShadow(value, label = "Presentation image shadow") {
  if (value == null || value === false || value === "none") return undefined;
  let source = typeof value === "string" ? SHADOW_PRESETS[value] : value;
  if (!source && typeof value === "string") {
    const match = /^(-?\d+(?:\.\d+)?)px\s+(-?\d+(?:\.\d+)?)px\s+(\d+(?:\.\d+)?)px\s+(#[0-9a-f]{6})(?:\/(\d+(?:\.\d+)?))?$/i.exec(value.trim());
    if (match) {
      const offsetX = Number(match[1]);
      const offsetY = Number(match[2]);
      source = {
        color: match[4],
        blurRadius: Number(match[3]),
        distance: Math.hypot(offsetX, offsetY),
        direction: (Math.atan2(offsetY, offsetX) * 180 / Math.PI + 360) % 360,
        opacity: match[5] == null ? 1 : Number(match[5]) / 100,
      };
    }
  }
  const input = assertObject(source, label);
  const unsupported = Object.keys(input).filter((key) => !SHADOW_KEYS.has(key));
  if (unsupported.length) throw new RangeError(`${label} uses unsupported properties: ${unsupported.sort().join(", ")}.`);
  const rawScheme = input.colorScheme ?? input.scheme;
  const color = input.color ?? input.fill;
  if (rawScheme != null && color != null && String(color).trim()) throw new RangeError(`${label} cannot combine color with colorScheme.`);
  const colorScheme = rawScheme == null ? undefined : String(rawScheme).trim().toLowerCase();
  if (colorScheme !== undefined && !SHADOW_SCHEME_COLORS.has(colorScheme)) throw new RangeError(`${label}.colorScheme must be a supported theme color token.`);
  const normalizedColor = colorScheme === undefined ? normalizeColor(color ?? "#000000", label) : undefined;
  const direction = ((boundedNumber(input.direction ?? input.angle ?? 0, `${label}.direction`, { min: -Infinity, max: Infinity }) % 360) + 360) % 360;
  const alignment = input.alignment == null ? undefined : String(input.alignment).trim();
  if (alignment !== undefined && !SHADOW_ALIGNMENTS.has(alignment)) throw new RangeError(`${label}.alignment must be a supported DrawingML rectangle alignment.`);
  const rotateWithShape = input.rotateWithShape == null ? undefined : input.rotateWithShape;
  if (rotateWithShape !== undefined && typeof rotateWithShape !== "boolean") throw new TypeError(`${label}.rotateWithShape must be boolean.`);
  return {
    ...(colorScheme === undefined ? { color: normalizedColor } : { colorScheme }),
    blurRadius: boundedNumber(input.blurRadius ?? input.blur ?? 0, `${label}.blurRadius`),
    distance: boundedNumber(input.distance ?? 0, `${label}.distance`),
    direction,
    opacity: boundedNumber(input.opacity ?? 0.2, `${label}.opacity`, { max: 1 }),
    ...(alignment === undefined ? {} : { alignment }),
    ...(rotateWithShape === undefined ? {} : { rotateWithShape }),
  };
}

export function normalizePresentationImageMask(value, label = "Presentation image mask") {
  if (value == null || value === false || value === "none" || value === "rect" || value === "rectangle") return undefined;
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a supported DrawingML preset name.`);
  const token = value.trim();
  const canonical = token === "circle" ? "ellipse" : token === "round-rect" || token === "roundRectangle" ? "roundRect" : token;
  if (!IMAGE_MASK_PRESETS.has(canonical)) {
    throw new RangeError(`${label} must use a supported DrawingML preset geometry.`);
  }
  return canonical === "rect" ? undefined : canonical;
}
