/**
 * Source-aware statistical formula semantics.
 *
 * The evaluator supplies already bounded argument views. This leaf owns only
 * coercion and numerically stable statistics; it has no workbook/model access.
 */

const SINGLE_SERIES_FUNCTIONS = new Set(["STDEV.S", "STDEV.P", "VAR.S", "VAR.P"]);
const PAIRWISE_FUNCTIONS = new Set(["CORREL", "COVARIANCE.S", "COVARIANCE.P"]);

export const STATISTICS_SPILL_RANGE_FUNCTIONS = Object.freeze([
  ...SINGLE_SERIES_FUNCTIONS,
  ...PAIRWISE_FUNCTIONS,
]);

function directNumber(value, errorCode, numberText) {
  const error = errorCode(value);
  if (error) return { error };
  if (typeof value === "number") return Number.isFinite(value) ? { value } : { error: "#NUM!" };
  if (typeof value === "boolean") return { value: value ? 1 : 0 };
  if (typeof value === "string") {
    const number = numberText(value);
    return Number.isFinite(number) ? { value: number } : { error: "#VALUE!" };
  }
  return { error: "#VALUE!" };
}

function collectSeries(arguments_, helpers) {
  const numbers = [];
  for (const argument of arguments_) {
    for (const value of argument.values) {
      const error = helpers.errorCode(value);
      if (error && argument.source !== "reference") return { error, numbers: [] };
      if (argument.source === "reference") {
        if (error) continue;
        const number = helpers.referenceNumber(value);
        const numberError = helpers.errorCode(number);
        if (numberError) return { error: numberError, numbers: [] };
        if (number !== undefined) numbers.push(number);
        continue;
      }
      const direct = directNumber(value, helpers.errorCode, helpers.numberText);
      if (direct.error) return { error: direct.error, numbers: [] };
      numbers.push(direct.value);
    }
  }
  return { numbers };
}

function variance(numbers, sample) {
  if (numbers.length < (sample ? 2 : 1)) return "#DIV/0!";
  let count = 0;
  let mean = 0;
  let squaredDeviation = 0;
  for (const number of numbers) {
    count += 1;
    const delta = number - mean;
    mean += delta / count;
    squaredDeviation += delta * (number - mean);
  }
  const result = squaredDeviation / (sample ? count - 1 : count);
  return Number.isFinite(result) ? Math.max(0, result) : "#NUM!";
}

function collectPairs(left, right, helpers) {
  if (left.values.length !== right.values.length) return { error: "#N/A", pairs: [] };
  const pairs = [];
  for (let index = 0; index < left.values.length; index += 1) {
    const leftValue = left.values[index];
    const rightValue = right.values[index];
    const error = helpers.errorCode(leftValue) || helpers.errorCode(rightValue);
    if (error) return { error, pairs: [] };
    const leftNumber = helpers.referenceNumber(leftValue);
    const rightNumber = helpers.referenceNumber(rightValue);
    const numberError = helpers.errorCode(leftNumber) || helpers.errorCode(rightNumber);
    if (numberError) return { error: numberError, pairs: [] };
    if (leftNumber === undefined || rightNumber === undefined) continue;
    pairs.push([leftNumber, rightNumber]);
  }
  return { pairs };
}

function pairMoments(pairs) {
  let count = 0;
  let meanX = 0;
  let meanY = 0;
  let coMoment = 0;
  let squaredDeviationX = 0;
  let squaredDeviationY = 0;
  for (const [x, y] of pairs) {
    count += 1;
    const deltaX = x - meanX;
    const deltaY = y - meanY;
    meanX += deltaX / count;
    meanY += deltaY / count;
    coMoment += deltaX * (y - meanY);
    squaredDeviationX += deltaX * (x - meanX);
    squaredDeviationY += deltaY * (y - meanY);
  }
  return { count, coMoment, squaredDeviationX, squaredDeviationY };
}

export function evaluateStatisticalFormula(fnName, args, helpers) {
  if (SINGLE_SERIES_FUNCTIONS.has(fnName)) {
    if (args.length < 1 || args.length > 254 || helpers.hasEmptyArgument()) return "#VALUE!";
    const collected = collectSeries(args.map((_, index) => helpers.argument(index)), helpers);
    if (collected.error) return collected.error;
    const sample = fnName.endsWith(".S");
    const result = variance(collected.numbers, sample);
    if (typeof result !== "number") return result;
    if (fnName.startsWith("STDEV")) {
      const standardDeviation = Math.sqrt(result);
      return Number.isFinite(standardDeviation) ? standardDeviation : "#NUM!";
    }
    return result;
  }

  if (PAIRWISE_FUNCTIONS.has(fnName)) {
    if (args.length !== 2 || helpers.hasEmptyArgument()) return "#VALUE!";
    const collected = collectPairs(helpers.argument(0), helpers.argument(1), helpers);
    if (collected.error) return collected.error;
    const moments = pairMoments(collected.pairs);
    if (fnName === "COVARIANCE.P") {
      if (moments.count < 1) return "#DIV/0!";
      const result = moments.coMoment / moments.count;
      return Number.isFinite(result) ? result : "#NUM!";
    }
    if (moments.count < 2) return "#DIV/0!";
    if (fnName === "COVARIANCE.S") {
      const result = moments.coMoment / (moments.count - 1);
      return Number.isFinite(result) ? result : "#NUM!";
    }
    if (moments.squaredDeviationX <= 0 || moments.squaredDeviationY <= 0) return "#DIV/0!";
    const result = moments.coMoment / Math.sqrt(moments.squaredDeviationX * moments.squaredDeviationY);
    return Number.isFinite(result) ? result : "#NUM!";
  }

  return undefined;
}
