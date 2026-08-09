const MAX_COORDINATE = 2_147_483_647;
const MAX_ADJUSTMENTS = 256;
const MAX_GUIDES = 1_024;
const MAX_FORMULA_LENGTH = 256;
const ANGLE_UNITS_PER_DEGREE = 60_000;
const GUIDE_NAME = /^[A-Za-z_][A-Za-z0-9_.-]{0,127}$/;
const INTEGER_TOKEN = /^-?\d+$/;

const FORMULA_ARITY = new Map([
  ["*/", 3],
  ["+-", 3],
  ["+/", 3],
  ["?:", 3],
  ["abs", 1],
  ["at2", 2],
  ["cat2", 3],
  ["cos", 2],
  ["max", 2],
  ["min", 2],
  ["mod", 3],
  ["pin", 3],
  ["sat2", 3],
  ["sin", 2],
  ["sqrt", 1],
  ["tan", 2],
  ["val", 1],
]);

const BUILTIN_NAMES = Object.freeze([
  "3cd4", "3cd8", "5cd8", "7cd8", "b", "cd2", "cd4", "cd8", "h", "hc",
  "hd2", "hd3", "hd4", "hd5", "hd6", "hd8", "l", "ls", "r", "ss",
  "ssd2", "ssd4", "ssd6", "ssd8", "ssd16", "ssd32", "t", "vc", "w",
  "wd2", "wd3", "wd4", "wd5", "wd6", "wd8", "wd10",
]);

function finiteBounded(value, label) {
  if (!Number.isFinite(value) || Math.abs(value) > MAX_COORDINATE) {
    throw new RangeError(`${label} must evaluate to a finite value within the DrawingML signed 32-bit range.`);
  }
  return Object.is(value, -0) ? 0 : value;
}

function guideName(value, label) {
  if (typeof value !== "string" || !GUIDE_NAME.test(value) || INTEGER_TOKEN.test(value)) {
    throw new TypeError(`${label} must be an ASCII DrawingML guide name without whitespace.`);
  }
  if (/^officeKit/i.test(value)) throw new TypeError(`${label} uses the reserved officeKit prefix.`);
  return value;
}

function formulaOperand(token, available, label) {
  if (INTEGER_TOKEN.test(token)) {
    const value = Number(token);
    if (!Number.isSafeInteger(value) || value < -MAX_COORDINATE || value > MAX_COORDINATE) {
      throw new RangeError(`${label} literal must be within the DrawingML signed 32-bit range.`);
    }
    return;
  }
  if (!available.has(token)) {
    throw new ReferenceError(`${label} references unknown or forward guide ${token}.`);
  }
}

function normalizedFormula(value, available, label) {
  if (typeof value !== "string") throw new TypeError(`${label} must be a DrawingML formula string.`);
  const formula = value.trim().split(/\s+/u).join(" ");
  if (!formula || formula.length > MAX_FORMULA_LENGTH) {
    throw new RangeError(`${label} must contain 1 through ${MAX_FORMULA_LENGTH} characters.`);
  }
  const tokens = formula.split(" ");
  const arity = FORMULA_ARITY.get(tokens[0]);
  if (arity === undefined) throw new TypeError(`${label} uses unsupported operator ${tokens[0]}.`);
  if (tokens.length !== arity + 1) throw new TypeError(`${label} operator ${tokens[0]} requires ${arity} operands.`);
  tokens.slice(1).forEach((token, index) => formulaOperand(token, available, `${label} operand ${index + 1}`));
  return formula;
}

function normalizeGuideList(value, label, limit, available) {
  if (value == null) return [];
  if (!Array.isArray(value) || value.length > limit) throw new RangeError(`${label} must contain at most ${limit} guides.`);
  return value.map((item, index) => {
    const itemLabel = `${label} ${index + 1}`;
    if (!item || typeof item !== "object" || Array.isArray(item)) throw new TypeError(`${itemLabel} must be an object.`);
    const unknown = Object.keys(item).filter((key) => key !== "name" && key !== "formula");
    if (unknown.length) throw new TypeError(`${itemLabel} has unsupported fields: ${unknown.join(", ")}.`);
    const name = guideName(item.name, `${itemLabel}.name`);
    if (available.has(name)) throw new TypeError(`${itemLabel}.name duplicates built-in or earlier guide ${name}.`);
    const formula = normalizedFormula(item.formula, available, `${itemLabel}.formula`);
    available.add(name);
    return { name, formula };
  });
}

export function normalizePresentationCustomGeometryFormulaGraph(value = {}) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError("Presentation custom geometry formula graph must be an object.");
  const unknown = Object.keys(value).filter((key) => key !== "adjustments" && key !== "guides");
  if (unknown.length) throw new TypeError(`Presentation custom geometry formula graph has unsupported fields: ${unknown.join(", ")}.`);
  const available = new Set(BUILTIN_NAMES);
  const adjustments = normalizeGuideList(value.adjustments, "Presentation custom adjustment", MAX_ADJUSTMENTS, available);
  const guides = normalizeGuideList(value.guides, "Presentation custom guide", MAX_GUIDES, available);
  return { adjustments, guides };
}

export function presentationCustomGeometryReferenceNames(graph = {}) {
  const normalized = normalizePresentationCustomGeometryFormulaGraph(graph);
  return new Set([...normalized.adjustments.map((item) => item.name), ...normalized.guides.map((item) => item.name)]);
}

export function normalizePresentationCustomGeometryReference(value, references, label) {
  if (typeof value !== "string" || INTEGER_TOKEN.test(value) || !references.has(value)) {
    throw new TypeError(`${label} must be a number or a declared DrawingML guide reference.`);
  }
  return value;
}

function builtins(widthEmu, heightEmu) {
  const w = finiteBounded(Number(widthEmu), "Presentation custom geometry width");
  const h = finiteBounded(Number(heightEmu), "Presentation custom geometry height");
  if (w <= 0 || h <= 0) throw new RangeError("Presentation custom geometry formula evaluation requires positive shape extents.");
  const ss = Math.min(w, h);
  return new Map([
    ["3cd4", 16_200_000], ["3cd8", 8_100_000], ["5cd8", 13_500_000], ["7cd8", 18_900_000],
    ["b", h], ["cd2", 10_800_000], ["cd4", 5_400_000], ["cd8", 2_700_000], ["h", h],
    ["hc", w / 2], ["hd2", h / 2], ["hd3", h / 3], ["hd4", h / 4], ["hd5", h / 5],
    ["hd6", h / 6], ["hd8", h / 8], ["l", 0], ["ls", Math.max(w, h)], ["r", w], ["ss", ss],
    ["ssd2", ss / 2], ["ssd4", ss / 4], ["ssd6", ss / 6], ["ssd8", ss / 8],
    ["ssd16", ss / 16], ["ssd32", ss / 32], ["t", 0], ["vc", h / 2], ["w", w],
    ["wd2", w / 2], ["wd3", w / 3], ["wd4", w / 4], ["wd5", w / 5], ["wd6", w / 6],
    ["wd8", w / 8], ["wd10", w / 10],
  ]);
}

function operand(token, values, label) {
  if (INTEGER_TOKEN.test(token)) return Number(token);
  if (!values.has(token)) throw new ReferenceError(`${label} references unavailable guide ${token}.`);
  return values.get(token);
}

function angleRadians(value) {
  return value / ANGLE_UNITS_PER_DEGREE * Math.PI / 180;
}

function evaluateFormula(formula, values, label) {
  const [operator, ...tokens] = formula.split(" ");
  const args = tokens.map((token) => operand(token, values, label));
  let result;
  switch (operator) {
    case "*/":
      if (args[2] === 0) throw new RangeError(`${label} divides by zero.`);
      result = args[0] * args[1] / args[2];
      break;
    case "+-": result = args[0] + args[1] - args[2]; break;
    case "+/":
      if (args[2] === 0) throw new RangeError(`${label} divides by zero.`);
      result = (args[0] + args[1]) / args[2];
      break;
    case "?:": result = args[0] > 0 ? args[1] : args[2]; break;
    case "abs": result = Math.abs(args[0]); break;
    case "at2": result = Math.atan2(args[1], args[0]) * 180 / Math.PI * ANGLE_UNITS_PER_DEGREE; break;
    case "cat2": result = args[0] * Math.cos(Math.atan2(args[2], args[1])); break;
    case "cos": result = args[0] * Math.cos(angleRadians(args[1])); break;
    case "max": result = Math.max(args[0], args[1]); break;
    case "min": result = Math.min(args[0], args[1]); break;
    case "mod": result = Math.hypot(args[0], args[1], args[2]); break;
    case "pin": result = args[1] < args[0] ? args[0] : args[1] > args[2] ? args[2] : args[1]; break;
    case "sat2": result = args[0] * Math.sin(Math.atan2(args[2], args[1])); break;
    case "sin": result = args[0] * Math.sin(angleRadians(args[1])); break;
    case "sqrt":
      if (args[0] < 0) throw new RangeError(`${label} takes the square root of a negative value.`);
      result = Math.sqrt(args[0]);
      break;
    case "tan": result = args[0] * Math.tan(angleRadians(args[1])); break;
    case "val": result = args[0]; break;
    default: throw new TypeError(`${label} uses unsupported operator ${operator}.`);
  }
  return finiteBounded(result, label);
}

export function evaluatePresentationCustomGeometryFormulaGraph(graph, { widthEmu, heightEmu } = {}) {
  const normalized = normalizePresentationCustomGeometryFormulaGraph(graph);
  const values = builtins(widthEmu, heightEmu);
  for (const [kind, items] of [["adjustment", normalized.adjustments], ["guide", normalized.guides]]) {
    for (const item of items) values.set(item.name, evaluateFormula(item.formula, values, `Presentation custom ${kind} ${item.name}`));
  }
  return values;
}

export function resolvePresentationCustomGeometryReference(value, values, label) {
  if (typeof value === "number") return finiteBounded(value, label);
  if (typeof value !== "string" || !values.has(value)) throw new ReferenceError(`${label} references unavailable guide ${value}.`);
  return finiteBounded(values.get(value), label);
}
