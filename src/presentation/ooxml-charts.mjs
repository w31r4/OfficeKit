import { normalizeChartTrendlines } from "../shared/chart-trendlines.mjs";

const BAR_GROUPINGS = new Set(["clustered", "stacked", "percentStacked"]);
const LINE_GROUPINGS = new Set(["standard", "stacked", "percentStacked"]);
const MARKER_SYMBOLS = new Set(["auto", "circle", "dash", "diamond", "dot", "none", "plus", "square", "star", "triangle", "x"]);
const DATA_LABEL_POSITIONS = new Set(["bestFit", "b", "ctr", "inBase", "inEnd", "l", "outEnd", "r", "t"]);
const DATA_LABEL_POSITION_ALIASES = new Map([
  ["bottom", "b"], ["center", "ctr"], ["insideBase", "inBase"], ["insideEnd", "inEnd"],
  ["left", "l"], ["outsideEnd", "outEnd"], ["right", "r"], ["top", "t"],
]);
const ERROR_BAR_DIRECTIONS = new Set(["x", "y"]);
const ERROR_BAR_TYPES = new Set(["both", "minus", "plus"]);
const ERROR_BAR_VALUE_TYPES = new Set(["cust", "fixedVal", "percentage", "stdDev", "stdErr"]);
const ERROR_BAR_VALUE_TYPE_ALIASES = new Map([["custom", "cust"], ["fixed", "fixedVal"], ["percent", "percentage"], ["standardDeviation", "stdDev"], ["standardError", "stdErr"]]);
const AXIS_GROUPS = new Set(["primary", "secondary"]);
const LINE_DASH_STYLES = new Set(["solid", "dot", "dash", "longDash", "dashDot", "longDashDot", "longDashDotDot", "systemDash", "systemDot", "systemDashDot", "systemDashDotDot"]);

function boundedInteger(value, { name, min, max, fallback, optional = false }) {
  if (value == null || value === "") return optional ? undefined : fallback;
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < min || parsed > max) throw new RangeError(`${name} must be an integer from ${min} to ${max}.`);
  return parsed;
}

function boundedNumber(value, { name, min, max, fallback, optional = false }) {
  if (value == null || value === "") return optional ? undefined : fallback;
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < min || parsed > max) throw new RangeError(`${name} must be a number from ${min} to ${max}.`);
  return parsed;
}

function enumValue(value, allowed, fallback, name) {
  if (value == null || value === "") return fallback;
  if (!allowed.has(value)) throw new TypeError(`${name} must be one of: ${[...allowed].join(", ")}.`);
  return value;
}

export function normalizePresentationChartMarker(marker) {
  if (marker == null || marker === false) return undefined;
  const raw = typeof marker === "string" ? { symbol: marker } : marker;
  if (!raw || typeof raw !== "object") throw new TypeError("chart marker must be a symbol string or object.");
  const fill = normalizePresentationChartPaint(raw.fill ?? raw.color);
  return {
    symbol: enumValue(raw.symbol || raw.style, MARKER_SYMBOLS, "auto", "chart marker symbol"),
    size: boundedInteger(raw.size, { name: "chart marker size", min: 2, max: 72, fallback: 5 }),
    ...(fill ? { fill } : {}),
    ...(raw.line == null && raw.stroke == null ? {} : { line: normalizePresentationChartLine(raw.line ?? raw.stroke) }),
  };
}

function normalizePresentationChartPaint(value) {
  if (typeof value === "string" && value) return value;
  if (!value || typeof value !== "object") return undefined;
  return [value.fill, value.color, value.rgb].find((candidate) => typeof candidate === "string" && candidate) || undefined;
}

function normalizePresentationChartLine(line) {
  if (line == null || line === false) return undefined;
  const raw = typeof line === "string" ? { fill: line } : line;
  if (!raw || typeof raw !== "object") throw new TypeError("chart line must be a color string or object.");
  const style = raw.style || raw.dash || "solid";
  if (!LINE_DASH_STYLES.has(style)) throw new TypeError(`chart line style must be one of: ${[...LINE_DASH_STYLES.keys()].join(", ")}.`);
  return {
    fill: normalizePresentationChartPaint(raw.fill ?? raw.color),
    width: boundedNumber(raw.width, { name: "chart line width", min: 0.1, max: 100, fallback: 1 }),
    style,
  };
}

function normalizePresentationChartPoints(points, valueCount) {
  const seen = new Set();
  return (points || []).map((point) => {
    if (!point || typeof point !== "object") throw new TypeError("chart points must be objects.");
    const rawIndex = point.idx ?? point.index;
    if (rawIndex == null) throw new TypeError("chart point idx is required.");
    const idx = boundedInteger(rawIndex, { name: "chart point idx", min: 0, max: 1_048_575 });
    if (valueCount != null && idx >= valueCount) throw new RangeError(`chart point idx ${idx} is outside the series value range.`);
    if (seen.has(idx)) throw new TypeError(`chart point idx ${idx} is duplicated.`);
    seen.add(idx);
    const fill = normalizePresentationChartPaint(point.fill ?? point.color);
    const line = normalizePresentationChartLine(point.line ?? point.stroke);
    return { idx, ...(fill ? { fill } : {}), ...(line ? { line } : {}) };
  });
}

export function normalizePresentationChartStyle(chartType, config = {}) {
  const type = String(chartType || config.chartType || "bar").toLowerCase();
  const style = config.style && typeof config.style === "object" ? config.style : {};
  const rawBar = config.barOptions || style.bar || {};
  const rawLine = config.lineOptions || style.line || {};
  const directionValue = rawBar.direction || rawBar.barDirection;
  const direction = directionValue === "horizontal" ? "bar" : directionValue === "vertical" ? "column" : directionValue;
  return {
    styleId: boundedInteger(config.styleId ?? config.styleIndex ?? style.id, { name: "chart styleId", min: 1, max: 48, optional: true }),
    varyColors: Boolean(config.varyColors ?? style.varyColors ?? ["pie", "doughnut"].includes(type)),
    barOptions: {
      direction: enumValue(direction, new Set(["column", "bar"]), "column", "chart bar direction"),
      grouping: enumValue(rawBar.grouping, BAR_GROUPINGS, "clustered", "chart bar grouping"),
      gapWidth: boundedInteger(rawBar.gapWidth, { name: "chart gapWidth", min: 0, max: 500, fallback: 150 }),
      overlap: boundedInteger(rawBar.overlap, { name: "chart overlap", min: -100, max: 100, fallback: 0 }),
    },
    lineOptions: {
      grouping: enumValue(rawLine.grouping, LINE_GROUPINGS, "standard", "chart line grouping"),
      marker: normalizePresentationChartMarker(rawLine.marker),
      smooth: Boolean(rawLine.smooth),
    },
  };
}

export function normalizePresentationChartSeriesStyle(series = {}, valueCount) {
  return {
    color: normalizePresentationChartPaint(series.color ?? series.fill),
    line: normalizePresentationChartLine(series.line ?? series.stroke),
    points: normalizePresentationChartPoints(series.points, valueCount),
    marker: normalizePresentationChartMarker(series.marker),
    smooth: series.smooth == null ? undefined : Boolean(series.smooth),
  };
}

export function normalizePresentationChartAxisGroup(value, chartType) {
  const axisGroup = value == null || value === "" ? "primary" : String(value);
  if (!AXIS_GROUPS.has(axisGroup)) throw new TypeError("chart series axisGroup must be primary or secondary.");
  if (axisGroup === "secondary" && !["bar", "line"].includes(chartType)) throw new TypeError("secondary chart axes are supported only for bar and line series.");
  return axisGroup;
}

export function normalizePresentationChartDataLabels(value) {
  if (value === true) return { showValue: true, showCategoryName: false, position: "bestFit" };
  if (value === false || value == null) return { showValue: false, showCategoryName: false, position: "bestFit" };
  if (typeof value !== "object") throw new TypeError("chart dataLabels must be a boolean or object.");
  const rawPosition = value.position || "bestFit";
  const position = DATA_LABEL_POSITION_ALIASES.get(rawPosition) || rawPosition;
  if (!DATA_LABEL_POSITIONS.has(position)) throw new TypeError(`chart data-label position must be one of: ${[...DATA_LABEL_POSITIONS].join(", ")}.`);
  return {
    showValue: Boolean(value.showValue),
    showCategoryName: Boolean(value.showCategoryName ?? value.showCategory),
    ...(value.showPercent == null ? {} : { showPercent: Boolean(value.showPercent) }),
    position,
  };
}

export function normalizePresentationChartTrendlines(value, valueCount, chartType) {
  return normalizeChartTrendlines(value, {
    valueCount,
    chartType,
    normalizeLine: normalizePresentationChartLine,
  });
}

function normalizePresentationErrorBarValues(values, valueCount, name) {
  if (!Array.isArray(values) || values.length === 0) throw new TypeError(`${name} must be a non-empty numeric array.`);
  if (valueCount != null && values.length !== valueCount) throw new RangeError(`${name} must contain exactly ${valueCount} values.`);
  return values.map((value) => boundedNumber(value, { name, min: 0, max: Number.MAX_SAFE_INTEGER }));
}

function normalizePresentationErrorBarText(value, name, maxLength) {
  if (value == null) return undefined;
  if (typeof value !== "string") throw new TypeError(`${name} must be a string.`);
  const normalized = value.trim();
  if (!normalized) throw new TypeError(`${name} must be non-empty.`);
  if (normalized.length > maxLength) throw new RangeError(`${name} must contain at most ${maxLength} characters.`);
  if (name.endsWith("Formula") && normalized.startsWith("=")) throw new TypeError(`${name} must omit the leading equals sign.`);
  if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(normalized)) throw new TypeError(`${name} contains unsupported control characters.`);
  return normalized;
}

function normalizePresentationErrorBarSide(value, side, valueCount, required) {
  if (!required) return {};
  const source = value[side];
  const objectSource = source && typeof source === "object" && !Array.isArray(source) ? source : {};
  const formula = normalizePresentationErrorBarText(value[`${side}Formula`] ?? value[`${side}Reference`] ?? objectSource.formula ?? objectSource.reference, `chart error-bar ${side}Formula`, 8_192);
  const rawValues = value[`${side}Values`] ?? (Array.isArray(source) ? source : objectSource.values ?? objectSource.cache);
  const values = rawValues == null
    ? formula ? undefined : normalizePresentationErrorBarValues(rawValues, valueCount, `chart error-bar ${side}Values`)
    : normalizePresentationErrorBarValues(rawValues, valueCount, `chart error-bar ${side}Values`);
  const formatCode = normalizePresentationErrorBarText(value[`${side}FormatCode`] ?? objectSource.formatCode, `chart error-bar ${side}FormatCode`, 255);
  if (formatCode && !values) throw new TypeError(`chart error-bar ${side}FormatCode requires cached ${side}Values.`);
  return {
    ...(values ? { [`${side}Values`]: values } : {}),
    ...(formula ? { [`${side}Formula`]: formula } : {}),
    ...(formatCode ? { [`${side}FormatCode`]: formatCode } : {}),
  };
}

export function normalizePresentationChartErrorBars(value, chartType, valueCount) {
  if (value == null || value === false) return undefined;
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError("chart errorBars must be an object.");
  if (chartType === "pie") throw new TypeError("chart errorBars are supported only for bar and line series.");
  const direction = value.direction || "y";
  if (!ERROR_BAR_DIRECTIONS.has(direction)) throw new TypeError(`chart error-bar direction must be one of: ${[...ERROR_BAR_DIRECTIONS].join(", ")}.`);
  const type = value.type || "both";
  if (!ERROR_BAR_TYPES.has(type)) throw new TypeError(`chart error-bar type must be one of: ${[...ERROR_BAR_TYPES].join(", ")}.`);
  const rawValueType = value.valueType || value.kind || "fixedVal";
  const valueType = ERROR_BAR_VALUE_TYPE_ALIASES.get(rawValueType) || rawValueType;
  if (!ERROR_BAR_VALUE_TYPES.has(valueType)) throw new TypeError(`chart error-bar valueType must be one of: ${[...ERROR_BAR_VALUE_TYPES].join(", ")}.`);
  const amount = ["cust", "stdErr"].includes(valueType) ? undefined : boundedNumber(value.value ?? value.amount, { name: "chart error-bar value", min: 0, max: Number.MAX_SAFE_INTEGER, fallback: valueType === "percentage" ? 5 : 1 });
  if (valueType === "stdErr" && (value.value ?? value.amount) != null) throw new TypeError("standard-error chart error bars do not accept a value.");
  const plus = valueType === "cust" ? normalizePresentationErrorBarSide(value, "plus", valueCount, type !== "minus") : {};
  const minus = valueType === "cust" ? normalizePresentationErrorBarSide(value, "minus", valueCount, type !== "plus") : {};
  const line = normalizePresentationChartLine(value.line ?? value.stroke);
  return {
    direction,
    type,
    valueType,
    ...(amount == null ? {} : { value: amount }),
    ...plus,
    ...minus,
    noEndCap: Boolean(value.noEndCap ?? value.endStyle === "none"),
    ...(line ? { line } : {}),
  };
}
