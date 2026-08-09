const CURVE_SAMPLE_COUNT = 65;
const SOLVER_EPSILON = 1e-12;
const TYPES = new Set(["exp", "linear", "log", "movingAvg", "poly", "power"]);
const TYPE_ALIASES = new Map([
  ["exponential", "exp"],
  ["logarithmic", "log"],
  ["movingAverage", "movingAvg"],
  ["polynomial", "poly"],
]);

function boundedInteger(value, { name, min, max, fallback }) {
  if (value == null || value === "") return fallback;
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < min || parsed > max) {
    throw new RangeError(`${name} must be an integer from ${min} to ${max}.`);
  }
  return parsed;
}

function boundedOptionalNumber(value, { name, min, max }) {
  if (value == null || value === "") return undefined;
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < min || parsed > max) {
    throw new RangeError(`${name} must be a number from ${min} to ${max}.`);
  }
  return parsed;
}

function normalizeTrendline(value, valueCount, normalizeLine) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new TypeError("chart trendlines must be objects.");
  }
  const rawType = value.type || "linear";
  const type = TYPE_ALIASES.get(rawType) || rawType;
  if (!TYPES.has(type)) throw new TypeError(`chart trendline type must be one of: ${[...TYPES].join(", ")}.`);
  if (value.order != null && type !== "poly") throw new TypeError("chart trendline order is supported only for polynomial trendlines.");
  if (value.period != null && type !== "movingAvg") throw new TypeError("chart trendline period is supported only for moving-average trendlines.");

  const order = type === "poly"
    ? boundedInteger(value.order, { name: "polynomial chart trendline order", min: 2, max: 6, fallback: 2 })
    : undefined;
  const periodMax = valueCount == null ? 255 : Math.min(255, valueCount - 1);
  if (type === "movingAvg" && periodMax < 2) throw new RangeError("moving-average chart trendlines require at least three series values.");
  const period = type === "movingAvg"
    ? boundedInteger(value.period, { name: "moving-average chart trendline period", min: 2, max: periodMax, fallback: 2 })
    : undefined;
  const extension = (candidate, name) => {
    const normalized = boundedOptionalNumber(candidate, { name, min: 0, max: 1_000_000 });
    if (normalized != null && Math.abs(normalized * 2 - Math.round(normalized * 2)) > 1e-9) {
      throw new RangeError(`${name} must use 0.5 increments for category charts.`);
    }
    return normalized;
  };
  const forward = extension(value.forward, "chart trendline forward");
  const backward = extension(value.backward, "chart trendline backward");
  const intercept = boundedOptionalNumber(value.intercept, {
    name: "chart trendline intercept",
    min: -Number.MAX_SAFE_INTEGER,
    max: Number.MAX_SAFE_INTEGER,
  });
  const name = value.name == null ? undefined : String(value.name);
  if (name != null && (name.length < 1 || name.length > 255 || /\p{Cc}/u.test(name))) {
    throw new RangeError("chart trendline name must contain 1 to 255 characters without controls.");
  }
  const lineValue = value.line ?? value.stroke;
  const line = lineValue == null ? undefined : normalizeLine(lineValue);
  return {
    type,
    ...(name ? { name } : {}),
    ...(order == null ? {} : { order }),
    ...(period == null ? {} : { period }),
    ...(forward == null ? {} : { forward }),
    ...(backward == null ? {} : { backward }),
    ...(intercept == null ? {} : { intercept }),
    displayEquation: Boolean(value.displayEquation ?? value.showEquation),
    displayRSquared: Boolean(value.displayRSquared ?? value.showRSquared),
    ...(line ? { line } : {}),
  };
}

export function normalizeChartTrendlines(value, { valueCount, chartType, normalizeLine } = {}) {
  if (value == null || value === false) return [];
  if (typeof normalizeLine !== "function") throw new TypeError("chart trendline normalization requires a line normalizer.");
  const items = Array.isArray(value) ? value : [value];
  if (items.length > 0 && !["bar", "line"].includes(chartType)) {
    throw new TypeError("chart trendlines are supported only for bar and line series.");
  }
  if (items.length > 16) throw new RangeError("chart series support at most 16 trendlines.");
  return items.map((item) => normalizeTrendline(item, valueCount, normalizeLine));
}

function finiteSeriesPoints(values = []) {
  return values.map((value, index) => ({ x: index + 1, y: Number(value) }))
    .filter((point) => Number.isFinite(point.y));
}

function solveLinearSystem(matrix, vector) {
  const size = vector.length;
  const augmented = matrix.map((row, index) => [...row, vector[index]]);
  for (let column = 0; column < size; column++) {
    let pivot = column;
    for (let row = column + 1; row < size; row++) {
      if (Math.abs(augmented[row][column]) > Math.abs(augmented[pivot][column])) pivot = row;
    }
    if (Math.abs(augmented[pivot][column]) <= SOLVER_EPSILON) return undefined;
    [augmented[column], augmented[pivot]] = [augmented[pivot], augmented[column]];
    const divisor = augmented[column][column];
    for (let index = column; index <= size; index++) augmented[column][index] /= divisor;
    for (let row = 0; row < size; row++) {
      if (row === column) continue;
      const factor = augmented[row][column];
      for (let index = column; index <= size; index++) augmented[row][index] -= factor * augmented[column][index];
    }
  }
  const result = augmented.map((row) => row[size]);
  return result.every(Number.isFinite) ? result : undefined;
}

function leastSquares(points, basisFunctions) {
  if (points.length < basisFunctions.length || basisFunctions.length === 0) return undefined;
  const matrix = basisFunctions.map((left) => basisFunctions.map((right) => points.reduce((sum, point) => sum + left(point.x) * right(point.x), 0)));
  const vector = basisFunctions.map((basis) => points.reduce((sum, point) => sum + basis(point.x) * point.y, 0));
  return solveLinearSystem(matrix, vector);
}

function polynomialPredictor(points, order, fixedIntercept) {
  const scale = Math.max(1, ...points.map((point) => Math.abs(point.x)));
  const firstPower = fixedIntercept == null ? 0 : 1;
  const powers = Array.from({ length: order - firstPower + 1 }, (_, index) => index + firstPower);
  const shifted = points.map((point) => ({ x: point.x / scale, y: point.y - (fixedIntercept ?? 0) }));
  const coefficients = leastSquares(shifted, powers.map((power) => (x) => x ** power));
  if (!coefficients) return undefined;
  return (x) => (fixedIntercept ?? 0) + coefficients.reduce((sum, coefficient, index) => sum + coefficient * (x / scale) ** powers[index], 0);
}

function transformedLinearPredictor(points, transformX, transformY, restoreY, transformedIntercept) {
  const transformed = [];
  for (const point of points) {
    const x = transformX(point.x);
    const y = transformY(point.y);
    if (Number.isFinite(x) && Number.isFinite(y)) transformed.push({ x, y });
  }
  if (transformed.length !== points.length || transformed.length < 2) return undefined;
  const shifted = transformedIntercept == null
    ? transformed
    : transformed.map((point) => ({ ...point, y: point.y - transformedIntercept }));
  const basis = transformedIntercept == null ? [() => 1, (x) => x] : [(x) => x];
  const coefficients = leastSquares(shifted, basis);
  if (!coefficients) return undefined;
  return (rawX) => {
    const x = transformX(rawX);
    if (!Number.isFinite(x)) return Number.NaN;
    const fitted = (transformedIntercept ?? 0) + coefficients.reduce((sum, coefficient, index) => sum + coefficient * basis[index](x), 0);
    return restoreY(fitted);
  };
}

function trendlinePredictor(trendline, points) {
  if (trendline.type === "linear") return polynomialPredictor(points, 1, trendline.intercept);
  if (trendline.type === "poly") return polynomialPredictor(points, trendline.order || 2, trendline.intercept);
  if (trendline.type === "exp") {
    if (trendline.intercept != null && !(trendline.intercept > 0)) return undefined;
    return transformedLinearPredictor(points, (x) => x, (y) => y > 0 ? Math.log(y) : Number.NaN, Math.exp, trendline.intercept == null ? undefined : Math.log(trendline.intercept));
  }
  if (trendline.type === "log") return transformedLinearPredictor(points, (x) => x > 0 ? Math.log(x) : Number.NaN, (y) => y, (y) => y, trendline.intercept);
  if (trendline.type === "power") {
    if (trendline.intercept != null && !(trendline.intercept > 0)) return undefined;
    return transformedLinearPredictor(points, (x) => x > 0 ? Math.log(x) : Number.NaN, (y) => y > 0 ? Math.log(y) : Number.NaN, Math.exp, trendline.intercept == null ? undefined : Math.log(trendline.intercept));
  }
  return undefined;
}

function movingAveragePoints(values, period) {
  const points = [];
  for (let end = period - 1; end < values.length; end++) {
    const window = values.slice(end - period + 1, end + 1).map(Number);
    if (!window.every(Number.isFinite)) continue;
    points.push({ x: end + 1, y: window.reduce((sum, value) => sum + value, 0) / period });
  }
  return points;
}

function trendlineDomain(trendline, categoryCount) {
  if (trendline.type === "movingAvg") return { start: 1, end: categoryCount };
  return {
    start: 1 - (trendline.backward || 0),
    end: categoryCount + (trendline.forward || 0),
  };
}

function curveSegments(predict, domain, sampleCount = CURVE_SAMPLE_COUNT) {
  const segments = [];
  let current = [];
  const count = Math.max(2, Math.min(257, Math.trunc(sampleCount) || CURVE_SAMPLE_COUNT));
  for (let index = 0; index < count; index++) {
    const x = domain.start + (index / (count - 1)) * (domain.end - domain.start);
    const point = { x, y: predict(x) };
    if (Number.isFinite(point.y)) current.push(point);
    else if (current.length) {
      if (current.length > 1) segments.push(current);
      current = [];
    }
  }
  if (current.length > 1) segments.push(current);
  return segments;
}

export function sampleChartTrendline(values, trendline, options = {}) {
  const categoryCount = Math.max(0, Math.trunc(options.categoryCount ?? values?.length ?? 0));
  if (categoryCount < 2) return [];
  const domain = trendlineDomain(trendline, categoryCount);
  if (!(domain.end > domain.start)) return [];
  if (trendline.type === "movingAvg") {
    const points = movingAveragePoints(values || [], trendline.period || 2);
    return points.length > 1 ? [{ domain, points }] : [];
  }
  const sourcePoints = finiteSeriesPoints(values);
  if (sourcePoints.length < 2) return [];
  const predict = trendlinePredictor(trendline, sourcePoints);
  return predict ? curveSegments(predict, domain, options.sampleCount).map((points) => ({ domain, points })) : [];
}
