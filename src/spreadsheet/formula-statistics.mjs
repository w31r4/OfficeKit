/**
 * Source-aware statistical formula semantics.
 *
 * The evaluator supplies already bounded argument views. This leaf owns only
 * coercion and numerically stable statistics; it has no workbook/model access.
 */

const SINGLE_SERIES_FUNCTIONS = new Set(["STDEV.S", "STDEV.P", "VAR.S", "VAR.P"]);
const PAIRWISE_FUNCTIONS = new Set(["CORREL", "COVARIANCE.S", "COVARIANCE.P"]);
const REGRESSION_FUNCTIONS = new Set(["SLOPE", "INTERCEPT", "RSQ", "STEYX"]);
const FORECAST_FUNCTIONS = new Set(["FORECAST.LINEAR"]);

export const STATISTICS_SPILL_RANGE_FUNCTIONS = Object.freeze([
  ...SINGLE_SERIES_FUNCTIONS,
  ...PAIRWISE_FUNCTIONS,
  ...REGRESSION_FUNCTIONS,
  ...FORECAST_FUNCTIONS,
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
    const leftDirect = left.source === "reference"
      ? undefined
      : directNumber(leftValue, helpers.errorCode, helpers.numberText);
    const rightDirect = right.source === "reference"
      ? undefined
      : directNumber(rightValue, helpers.errorCode, helpers.numberText);
    if (leftDirect?.error || rightDirect?.error) return { error: leftDirect?.error || rightDirect.error, pairs: [] };
    const leftNumber = left.source === "reference" ? helpers.referenceNumber(leftValue) : leftDirect.value;
    const rightNumber = right.source === "reference" ? helpers.referenceNumber(rightValue) : rightDirect.value;
    const numberError = helpers.errorCode(leftNumber) || helpers.errorCode(rightNumber);
    if (numberError) return { error: numberError, pairs: [] };
    if (leftNumber === undefined || rightNumber === undefined) continue;
    pairs.push([leftNumber, rightNumber]);
  }
  return { pairs };
}

function pairMoments(pairs) {
  let count = 0;
  let meanLeft = 0;
  let meanRight = 0;
  let coMoment = 0;
  let squaredDeviationLeft = 0;
  let squaredDeviationRight = 0;
  for (const [left, right] of pairs) {
    count += 1;
    const deltaLeft = left - meanLeft;
    const deltaRight = right - meanRight;
    meanLeft += deltaLeft / count;
    meanRight += deltaRight / count;
    coMoment += deltaLeft * (right - meanRight);
    squaredDeviationLeft += deltaLeft * (left - meanLeft);
    squaredDeviationRight += deltaRight * (right - meanRight);
  }
  return {
    count,
    meanLeft,
    meanRight,
    coMoment,
    squaredDeviationLeft,
    squaredDeviationRight,
  };
}

function regression(moments) {
  if (moments.squaredDeviationRight <= 0) return { error: "#DIV/0!" };
  const slope = moments.coMoment / moments.squaredDeviationRight;
  const intercept = moments.meanLeft - slope * moments.meanRight;
  return Number.isFinite(slope) && Number.isFinite(intercept)
    ? { slope, intercept }
    : { error: "#NUM!" };
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
    if (moments.squaredDeviationLeft <= 0 || moments.squaredDeviationRight <= 0) return "#DIV/0!";
    const result = moments.coMoment / Math.sqrt(moments.squaredDeviationLeft * moments.squaredDeviationRight);
    return Number.isFinite(result) ? result : "#NUM!";
  }

  if (REGRESSION_FUNCTIONS.has(fnName) || FORECAST_FUNCTIONS.has(fnName)) {
    const forecast = FORECAST_FUNCTIONS.has(fnName);
    if (args.length !== (forecast ? 3 : 2) || helpers.hasEmptyArgument()) return "#VALUE!";
    let predictor;
    if (forecast) {
      const argument = helpers.argument(0);
      if (argument.values.length !== 1) return "#VALUE!";
      const value = argument.values[0];
      const error = helpers.errorCode(value);
      if (error) return error;
      const numeric = argument.source === "reference"
        ? helpers.referenceNumber(value)
        : directNumber(value, helpers.errorCode, helpers.numberText);
      if (numeric?.error) return numeric.error;
      const predictorNumber = argument.source === "reference" ? numeric : numeric.value;
      if (helpers.errorCode(predictorNumber)) return predictorNumber;
      if (predictorNumber === undefined) return "#VALUE!";
      predictor = predictorNumber;
    }
    const offset = forecast ? 1 : 0;
    const collected = collectPairs(helpers.argument(offset), helpers.argument(offset + 1), helpers);
    if (collected.error) return collected.error;
    const moments = pairMoments(collected.pairs);
    if (moments.count === 0) return fnName === "STEYX" ? "#DIV/0!" : "#N/A";
    if (fnName === "RSQ") {
      if (moments.count < 2 || moments.squaredDeviationLeft <= 0 || moments.squaredDeviationRight <= 0) return "#DIV/0!";
      const result = (moments.coMoment * moments.coMoment)
        / (moments.squaredDeviationLeft * moments.squaredDeviationRight);
      return Number.isFinite(result) ? Math.min(1, Math.max(0, result)) : "#NUM!";
    }
    const fitted = regression(moments);
    if (fitted.error) return fitted.error;
    if (fnName === "SLOPE") return fitted.slope;
    if (fnName === "INTERCEPT") return fitted.intercept;
    if (fnName === "STEYX") {
      if (moments.count < 3) return "#DIV/0!";
      let residual = 0;
      for (const [left, right] of collected.pairs) {
        const fittedLeft = moments.meanLeft + fitted.slope * (right - moments.meanRight);
        const error = left - fittedLeft;
        residual += error * error;
      }
      const result = Math.sqrt(residual / (moments.count - 2));
      return Number.isFinite(result) ? result : "#NUM!";
    }
    // Evaluate the prediction around the observed x mean. This is algebraically
    // identical to intercept + slope*x but avoids needlessly cancelling two
    // large terms when the source domain has a large offset.
    const result = moments.meanLeft + fitted.slope * (predictor - moments.meanRight);
    return Number.isFinite(result) ? result : "#NUM!";
  }

  return undefined;
}
