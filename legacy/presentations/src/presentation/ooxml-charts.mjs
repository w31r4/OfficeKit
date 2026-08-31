import { normalizeChartTrendlines } from "../shared/chart-trendlines.mjs";
import { normalizeChartErrorBars } from "../shared/chart-error-bars.mjs";

const BAR_GROUPINGS = new Set(["clustered", "stacked", "percentStacked"]);
const LINE_GROUPINGS = new Set(["standard", "stacked", "percentStacked"]);
const MARKER_SYMBOLS = new Set(["auto", "circle", "dash", "diamond", "dot", "none", "plus", "square", "star", "triangle", "x"]);
const DATA_LABEL_POSITIONS = new Set(["bestFit", "b", "ctr", "inBase", "inEnd", "l", "outEnd", "r", "t"]);
const DATA_LABEL_POSITION_ALIASES = new Map([
  ["bottom", "b"], ["center", "ctr"], ["insideBase", "inBase"], ["insideEnd", "inEnd"],
  ["left", "l"], ["outsideEnd", "outEnd"], ["right", "r"], ["top", "t"],
]);
const AXIS_GROUPS = new Set(["primary", "secondary"]);
const LINE_DASH_STYLES = new Set(["solid", "dot", "dash", "longDash", "dashDot", "longDashDot", "longDashDotDot", "systemDash", "systemDot", "systemDashDot", "systemDashDotDot"]);
const LINE_CAPS = new Set(["flat", "round", "square"]);
const LINE_JOINS = new Set(["miter", "round", "bevel"]);

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

function normalizePresentationChartSurfaceFill(value, name) {
  if (value == null) return undefined;
  if (typeof value === "string") return { type: "solid", color: value, opacity: 1 };
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${name} must be a color string or a fill object.`);
  const type = String(value.type || (value.color || value.fill ? "solid" : "none"));
  if (type === "none") return { type: "none" };
  if (type !== "solid") throw new TypeError(`${name} supports only none or solid fills.`);
  const color = normalizePresentationChartPaint(value.color ?? value.fill);
  if (!color) throw new TypeError(`${name} solid fill requires a color.`);
  return {
    type: "solid",
    color,
    opacity: boundedNumber(value.opacity, { name: `${name} opacity`, min: 0, max: 1, fallback: 1 }),
  };
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
    ...(raw.opacity == null ? {} : { opacity: boundedNumber(raw.opacity, { name: "chart line opacity", min: 0, max: 1 }) }),
    ...(raw.cap == null ? {} : { cap: enumValue(raw.cap, LINE_CAPS, "flat", "chart line cap") }),
    ...(raw.join == null ? {} : { join: enumValue(raw.join, LINE_JOINS, "miter", "chart line join") }),
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
    chartAreaFill: normalizePresentationChartSurfaceFill(config.chartAreaFill ?? style.chartAreaFill, "chart area fill"),
    plotAreaFill: normalizePresentationChartSurfaceFill(config.plotAreaFill ?? style.plotAreaFill, "plot area fill"),
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

export function normalizePresentationChartErrorBars(value, chartType, valueCount) {
  return normalizeChartErrorBars(value, {
    valueCount,
    chartType,
    normalizeLine: normalizePresentationChartLine,
  });
}
