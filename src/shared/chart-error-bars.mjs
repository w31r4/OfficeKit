const DIRECTIONS = new Set(["x", "y"]);
const BAR_TYPES = new Set(["both", "minus", "plus"]);
const VALUE_TYPES = new Set(["cust", "fixedVal", "percentage", "stdDev", "stdErr"]);
const VALUE_TYPE_ALIASES = new Map([
  ["custom", "cust"],
  ["fixed", "fixedVal"],
  ["fixedValue", "fixedVal"],
  ["percent", "percentage"],
  ["standardDeviation", "stdDev"],
  ["standardError", "stdErr"],
]);
const REFERENCE_TYPES = new Set(["standardError", "percentage", "standardDeviation", "none"]);
const END_STYLES = new Set(["cap", "noCap", "none"]);
const ERROR_BAR_FIELDS = new Set([
  "direction", "type", "errorBarType", "barType", "valueType", "kind", "value", "amount",
  "endStyle", "noEndCap", "line", "stroke",
  "plus", "plusValues", "plusFormula", "plusReference", "plusFormatCode",
  "minus", "minusValues", "minusFormula", "minusReference", "minusFormatCode",
]);

function sameValue(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function oneAlias(values, name, normalize = (value) => value) {
  const present = values.filter((value) => value != null && value !== "").map(normalize);
  if (present.length === 0) return undefined;
  if (present.slice(1).some((value) => !sameValue(value, present[0]))) {
    throw new TypeError(`${name} aliases must describe the same value.`);
  }
  return present[0];
}

function boundedNumber(value, { name, min = 0, max = Number.MAX_SAFE_INTEGER }) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < min || parsed > max) {
    throw new RangeError(`${name} must be a number from ${min} to ${max}.`);
  }
  return parsed;
}

function boundedValues(values, valueCount, name) {
  if (!Array.isArray(values) || values.length === 0) throw new TypeError(`${name} must be a non-empty numeric array.`);
  if (valueCount != null && values.length !== valueCount) throw new RangeError(`${name} must contain exactly ${valueCount} values.`);
  return values.map((value) => boundedNumber(value, { name }));
}

function boundedText(value, name, maxLength) {
  if (value == null || value === "") return undefined;
  if (typeof value !== "string") throw new TypeError(`${name} must be a string.`);
  const output = value.trim();
  if (!output) throw new TypeError(`${name} must be non-empty.`);
  if (output.length > maxLength) throw new RangeError(`${name} must contain at most ${maxLength} characters.`);
  if (name.endsWith("Formula") && output.startsWith("=")) throw new TypeError(`${name} must omit the leading equals sign.`);
  if (/\p{Cc}/u.test(output)) throw new TypeError(`${name} contains unsupported control characters.`);
  return output;
}

function sideInput(value, side) {
  return [value[side], value[`${side}Values`], value[`${side}Formula`], value[`${side}Reference`], value[`${side}FormatCode`]]
    .some((candidate) => candidate != null);
}

function normalizeSide(value, side, valueCount, required) {
  if (!required) {
    if (sideInput(value, side)) throw new TypeError(`chart error-bar ${side} data is not valid when type excludes ${side}.`);
    return undefined;
  }
  const source = value[side];
  if (source != null && !Array.isArray(source) && (typeof source !== "object" || source === null)) {
    throw new TypeError(`chart error-bar ${side} must be an array or object.`);
  }
  const objectSource = source && typeof source === "object" && !Array.isArray(source) ? source : {};
  const unsupported = Object.keys(objectSource).filter((key) => !["formula", "reference", "values", "cache", "formatCode"].includes(key) && objectSource[key] != null);
  if (unsupported.length) throw new TypeError(`chart error-bar ${side} supports only formula, reference, values, cache, and formatCode; received ${unsupported.join(", ")}.`);

  const formula = oneAlias([
    value[`${side}Formula`],
    value[`${side}Reference`],
    objectSource.formula,
    objectSource.reference,
  ], `chart error-bar ${side}Formula`, (candidate) => boundedText(candidate, `chart error-bar ${side}Formula`, 8_192));
  const rawValues = oneAlias([
    value[`${side}Values`],
    Array.isArray(source) ? source : undefined,
    objectSource.values,
    objectSource.cache,
  ], `chart error-bar ${side}Values`, (candidate) => boundedValues(candidate, valueCount, `chart error-bar ${side}Values`));
  if (rawValues == null && formula == null) throw new TypeError(`chart error-bar ${side} requires literal values or a formula.`);
  const formatCode = oneAlias([
    value[`${side}FormatCode`],
    objectSource.formatCode,
  ], `chart error-bar ${side}FormatCode`, (candidate) => boundedText(candidate, `chart error-bar ${side}FormatCode`, 255));
  if (formatCode && rawValues == null) throw new TypeError(`chart error-bar ${side}FormatCode requires cached ${side}Values.`);
  return {
    ...(rawValues == null ? {} : { values: rawValues }),
    ...(formula == null ? {} : { formula }),
    ...(formatCode == null ? {} : { formatCode }),
  };
}

function normalizedValueType(value, referenceType) {
  const raw = oneAlias([
    value.valueType,
    value.kind,
    referenceType && referenceType !== "none" ? referenceType : undefined,
  ], "chart error-bar valueType", (candidate) => VALUE_TYPE_ALIASES.get(candidate) || candidate) || "fixedVal";
  if (!VALUE_TYPES.has(raw)) throw new TypeError(`chart error-bar valueType must be one of: ${[...VALUE_TYPES].join(", ")}.`);
  return raw;
}

function normalizedBarType(value) {
  const raw = oneAlias([
    BAR_TYPES.has(value.type) ? value.type : undefined,
    value.errorBarType,
    value.barType,
  ], "chart error-bar type") || "both";
  if (!BAR_TYPES.has(raw)) throw new TypeError(`chart error-bar type must be one of: ${[...BAR_TYPES].join(", ")}.`);
  return raw;
}

function normalizedNoEndCap(value) {
  const endStyle = value.endStyle;
  if (endStyle != null && !END_STYLES.has(endStyle)) throw new TypeError(`chart error-bar endStyle must be one of: ${[...END_STYLES].join(", ")}.`);
  const fromStyle = endStyle == null ? undefined : endStyle !== "cap";
  if (value.noEndCap != null && typeof value.noEndCap !== "boolean") throw new TypeError("chart error-bar noEndCap must be boolean.");
  const explicit = value.noEndCap;
  if (fromStyle != null && explicit != null && fromStyle !== explicit) throw new TypeError("chart error-bar noEndCap and endStyle aliases must describe the same value.");
  return explicit ?? fromStyle ?? false;
}

export function normalizeChartErrorBars(value, { valueCount, chartType, normalizeLine } = {}) {
  if (value == null || value === false) return undefined;
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError("chart errorBars must be an object.");
  const unsupported = Object.keys(value).filter((key) => !ERROR_BAR_FIELDS.has(key) && value[key] != null);
  if (unsupported.length) throw new TypeError(`chart errorBars received unsupported fields: ${unsupported.join(", ")}.`);
  if (!["bar", "line"].includes(chartType)) throw new TypeError("chart errorBars are supported only for bar and line series.");
  const referenceType = REFERENCE_TYPES.has(value.type) ? value.type : undefined;
  if (value.type === "none") return undefined;
  if (value.type != null && !BAR_TYPES.has(value.type) && !referenceType) {
    throw new TypeError(`chart error-bar type must be one of: ${[...BAR_TYPES, ...REFERENCE_TYPES].join(", ")}.`);
  }
  const direction = value.direction || "y";
  if (!DIRECTIONS.has(direction)) throw new TypeError(`chart error-bar direction must be one of: ${[...DIRECTIONS].join(", ")}.`);
  const type = normalizedBarType(value);
  const valueType = normalizedValueType(value, referenceType);
  const amount = oneAlias([value.value, value.amount], "chart error-bar value", (candidate) => boundedNumber(candidate, { name: "chart error-bar value" }));
  if (["cust", "stdErr"].includes(valueType) && amount != null) {
    throw new TypeError(`${valueType === "cust" ? "custom" : "standard-error"} chart error bars do not accept a value.`);
  }
  const normalizedAmount = ["cust", "stdErr"].includes(valueType)
    ? undefined
    : amount ?? (valueType === "percentage" ? 5 : 1);
  if (valueType !== "cust" && ["plus", "minus"].some((side) => sideInput(value, side))) {
    throw new TypeError("chart error-bar plus/minus data is valid only for custom error bars.");
  }
  const plus = valueType === "cust" ? normalizeSide(value, "plus", valueCount, type !== "minus") : undefined;
  const minus = valueType === "cust" ? normalizeSide(value, "minus", valueCount, type !== "plus") : undefined;
  if (value.line != null && value.stroke != null && !sameValue(value.line, value.stroke)) {
    throw new TypeError("chart error-bar line and stroke aliases must describe the same style.");
  }
  if ((value.line != null || value.stroke != null) && typeof normalizeLine !== "function") {
    throw new TypeError("chart error-bar normalization requires a line normalizer when line styling is present.");
  }
  const line = value.line == null && value.stroke == null ? undefined : normalizeLine(value.line ?? value.stroke);
  return {
    direction,
    type,
    valueType,
    ...(normalizedAmount == null ? {} : { value: normalizedAmount }),
    ...(plus?.values == null ? {} : { plusValues: plus.values }),
    ...(plus?.formula == null ? {} : { plusFormula: plus.formula }),
    ...(plus?.formatCode == null ? {} : { plusFormatCode: plus.formatCode }),
    ...(minus?.values == null ? {} : { minusValues: minus.values }),
    ...(minus?.formula == null ? {} : { minusFormula: minus.formula }),
    ...(minus?.formatCode == null ? {} : { minusFormatCode: minus.formatCode }),
    noEndCap: normalizedNoEndCap(value),
    ...(line == null ? {} : { line }),
  };
}

export function chartErrorBarMagnitudes(values, errorBars) {
  if (!errorBars) return [];
  const numeric = (values || []).map(Number);
  const finite = numeric.filter(Number.isFinite);
  const mean = finite.reduce((sum, value) => sum + value, 0) / Math.max(1, finite.length);
  const deviation = Math.sqrt(finite.reduce((sum, value) => sum + (value - mean) ** 2, 0) / Math.max(1, finite.length));
  const magnitude = (value, index, side) => {
    if (errorBars.valueType === "cust") return Number(errorBars[`${side}Values`]?.[index]) || 0;
    if (errorBars.valueType === "percentage") return Math.abs(Number(value) || 0) * (errorBars.value || 0) / 100;
    if (errorBars.valueType === "stdDev") return deviation * (errorBars.value ?? 1);
    if (errorBars.valueType === "stdErr") return deviation / Math.sqrt(Math.max(1, finite.length));
    return errorBars.value || 0;
  };
  return numeric.map((value, index) => ({
    minus: errorBars.type === "plus" ? 0 : magnitude(value, index, "minus"),
    plus: errorBars.type === "minus" ? 0 : magnitude(value, index, "plus"),
  }));
}
