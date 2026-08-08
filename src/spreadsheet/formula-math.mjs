// Pure bounded math/trig formula semantics. The evaluator supplies already
// resolved scalar/range views so this module cannot reach back into workbook
// state or accidentally create a second reference engine.

const MATH_SPILL_RANGE_FUNCTIONS = new Set(["GCD", "LCM"]);

function strictFormulaNumber(value, formulaErrorCode) {
  const error = formulaErrorCode(value);
  if (error) return error;
  if (value == null || (typeof value === "string" && value.trim() === "")) return "#VALUE!";
  const number = Number(value);
  if (Number.isNaN(number)) return "#VALUE!";
  return Number.isFinite(number) ? number : "#NUM!";
}

function strictFormulaInteger(value, formulaErrorCode) {
  const number = strictFormulaNumber(value, formulaErrorCode);
  if (formulaErrorCode(number)) return number;
  const integer = Math.trunc(number);
  return Number.isSafeInteger(integer) ? integer : "#NUM!";
}

function formulaGcd(left, right) {
  let a = Math.abs(left);
  let b = Math.abs(right);
  while (b !== 0) {
    const remainder = a % b;
    a = b;
    b = remainder;
  }
  return a;
}

function formulaBinomial(n, k) {
  if (k < 0 || k > n) return "#NUM!";
  let result = 1;
  const reduced = Math.min(k, n - k);
  for (let index = 1; index <= reduced; index += 1) {
    result = result * (n - reduced + index) / index;
    if (!Number.isFinite(result)) return "#NUM!";
  }
  return Math.round(result);
}

function formulaFactorial(number, step = 1) {
  if (number < 0 || number > 170) return "#NUM!";
  let result = 1;
  for (let value = number; value > 1; value -= step) {
    result *= value;
    if (!Number.isFinite(result)) return "#NUM!";
  }
  return result;
}

function formulaRoundToParity(number, parity) {
  const magnitude = Math.ceil(Math.abs(number));
  const rounded = magnitude % 2 === parity ? magnitude : magnitude + 1;
  return number < 0 ? -rounded : rounded;
}

function evaluateFiniteUnaryMath(args, scalar, hasEmptyArgument, formulaErrorCode, operation) {
  if (args.length !== 1 || hasEmptyArgument()) return "#VALUE!";
  const number = strictFormulaNumber(scalar(0), formulaErrorCode);
  if (formulaErrorCode(number)) return number;
  const result = operation(number);
  return Number.isFinite(result) ? result : "#NUM!";
}

function evaluateLogFormula(args, scalar, hasEmptyArgument, formulaErrorCode) {
  if (args.length < 1 || args.length > 2 || hasEmptyArgument()) return "#VALUE!";
  const number = strictFormulaNumber(scalar(0), formulaErrorCode);
  const base = strictFormulaNumber(scalar(1, 10), formulaErrorCode);
  if (formulaErrorCode(number)) return number;
  if (formulaErrorCode(base)) return base;
  if (number <= 0 || base <= 0 || base === 1) return "#NUM!";
  const result = Math.log(number) / Math.log(base);
  return Number.isFinite(result) ? result : "#NUM!";
}

function evaluateAtan2Formula(args, scalar, hasEmptyArgument, formulaErrorCode) {
  if (args.length !== 2 || hasEmptyArgument()) return "#VALUE!";
  const x = strictFormulaNumber(scalar(0), formulaErrorCode);
  const y = strictFormulaNumber(scalar(1), formulaErrorCode);
  if (formulaErrorCode(x)) return x;
  if (formulaErrorCode(y)) return y;
  if (x === 0 && y === 0) return "#DIV/0!";
  const result = Math.atan2(y, x);
  return Number.isFinite(result) ? result : "#NUM!";
}

function evaluateMathFormula(fnName, args, { scalar, values, hasEmptyArgument, formulaErrorCode }) {
  switch (fnName) {
    case "GCD":
    case "LCM": {
      if (args.length === 0 || args.length > 255 || hasEmptyArgument()) return "#VALUE!";
      const integers = values().flatMap((value) => value == null || value === "" ? [] : [strictFormulaInteger(value, formulaErrorCode)]);
      const error = integers.map(formulaErrorCode).find(Boolean);
      if (error) return error;
      if (integers.length === 0) return 0;
      let result = Math.abs(integers[0]);
      for (const integer of integers.slice(1)) {
        const divisor = Math.abs(integer);
        if (fnName === "GCD") {
          result = formulaGcd(result, divisor);
          continue;
        }
        if (result === 0 || divisor === 0) {
          result = 0;
          continue;
        }
        const next = result / formulaGcd(result, divisor) * divisor;
        if (!Number.isSafeInteger(next)) return "#NUM!";
        result = next;
      }
      return result;
    }
    case "FACT":
    case "FACTDOUBLE": {
      if (args.length !== 1 || hasEmptyArgument()) return "#VALUE!";
      const number = strictFormulaInteger(scalar(0), formulaErrorCode);
      if (formulaErrorCode(number)) return number;
      return formulaFactorial(number, fnName === "FACTDOUBLE" ? 2 : 1);
    }
    case "COMBIN":
    case "COMBINA": {
      if (args.length !== 2 || hasEmptyArgument()) return "#VALUE!";
      const number = strictFormulaInteger(scalar(0), formulaErrorCode);
      const chosen = strictFormulaInteger(scalar(1), formulaErrorCode);
      if (formulaErrorCode(number)) return number;
      if (formulaErrorCode(chosen)) return chosen;
      if (number < 0 || chosen < 0) return "#NUM!";
      if (fnName === "COMBINA" && chosen === 0) return 1;
      const population = fnName === "COMBINA" ? number + chosen - 1 : number;
      return formulaBinomial(population, chosen);
    }
    case "MROUND": {
      if (args.length !== 2 || hasEmptyArgument()) return "#VALUE!";
      const number = strictFormulaNumber(scalar(0), formulaErrorCode);
      const multiple = strictFormulaNumber(scalar(1), formulaErrorCode);
      if (formulaErrorCode(number)) return number;
      if (formulaErrorCode(multiple)) return multiple;
      if (multiple === 0) return "#DIV/0!";
      if (number !== 0 && Math.sign(number) !== Math.sign(multiple)) return "#NUM!";
      const quotient = number / multiple;
      if (!Number.isFinite(quotient)) return "#NUM!";
      const rounded = Math.floor(Math.abs(quotient) + 0.5) * Math.abs(multiple);
      return Number.isFinite(rounded) ? (number < 0 ? -rounded : rounded) : "#NUM!";
    }
    case "EVEN":
    case "ODD": {
      if (args.length !== 1 || hasEmptyArgument()) return "#VALUE!";
      const number = strictFormulaNumber(scalar(0), formulaErrorCode);
      if (formulaErrorCode(number)) return number;
      const result = formulaRoundToParity(number, fnName === "EVEN" ? 0 : 1);
      return Number.isSafeInteger(result) ? result : "#NUM!";
    }
    case "EXP": return evaluateFiniteUnaryMath(args, scalar, hasEmptyArgument, formulaErrorCode, Math.exp);
    case "LN": return evaluateFiniteUnaryMath(args, scalar, hasEmptyArgument, formulaErrorCode, (number) => number > 0 ? Math.log(number) : NaN);
    case "LOG": return evaluateLogFormula(args, scalar, hasEmptyArgument, formulaErrorCode);
    case "LOG10": return evaluateFiniteUnaryMath(args, scalar, hasEmptyArgument, formulaErrorCode, (number) => number > 0 ? Math.log10(number) : NaN);
    case "SIN": return evaluateFiniteUnaryMath(args, scalar, hasEmptyArgument, formulaErrorCode, Math.sin);
    case "COS": return evaluateFiniteUnaryMath(args, scalar, hasEmptyArgument, formulaErrorCode, Math.cos);
    case "TAN": return evaluateFiniteUnaryMath(args, scalar, hasEmptyArgument, formulaErrorCode, Math.tan);
    case "ASIN": return evaluateFiniteUnaryMath(args, scalar, hasEmptyArgument, formulaErrorCode, (number) => number >= -1 && number <= 1 ? Math.asin(number) : NaN);
    case "ACOS": return evaluateFiniteUnaryMath(args, scalar, hasEmptyArgument, formulaErrorCode, (number) => number >= -1 && number <= 1 ? Math.acos(number) : NaN);
    case "ATAN": return evaluateFiniteUnaryMath(args, scalar, hasEmptyArgument, formulaErrorCode, Math.atan);
    case "ATAN2": return evaluateAtan2Formula(args, scalar, hasEmptyArgument, formulaErrorCode);
    case "SINH": return evaluateFiniteUnaryMath(args, scalar, hasEmptyArgument, formulaErrorCode, Math.sinh);
    case "COSH": return evaluateFiniteUnaryMath(args, scalar, hasEmptyArgument, formulaErrorCode, Math.cosh);
    case "TANH": return evaluateFiniteUnaryMath(args, scalar, hasEmptyArgument, formulaErrorCode, Math.tanh);
    case "ASINH": return evaluateFiniteUnaryMath(args, scalar, hasEmptyArgument, formulaErrorCode, Math.asinh);
    case "ACOSH": return evaluateFiniteUnaryMath(args, scalar, hasEmptyArgument, formulaErrorCode, (number) => number >= 1 ? Math.acosh(number) : NaN);
    case "ATANH": return evaluateFiniteUnaryMath(args, scalar, hasEmptyArgument, formulaErrorCode, (number) => number > -1 && number < 1 ? Math.atanh(number) : NaN);
    default:
      return undefined;
  }
}

export { MATH_SPILL_RANGE_FUNCTIONS, evaluateMathFormula };
