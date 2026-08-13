import { modelLetFormulaFromXlsx, xlsxLetFormulaFromModel } from "./office-kit-spreadsheet-let-syntax.mjs";

// Intersection of OfficeKit's modeled formula catalog and the future-function
// names in MS-XLSX section 2.2.3. Keep this package grammar out of the public
// worksheet model: Agents use the names Excel displays, while the codec writes
// the names Excel persists.
const XLFN_FUNCTIONS = [
  "CHOOSECOLS",
  "CHOOSEROWS",
  "COMBINA",
  "CONCAT",
  "COVARIANCE.P",
  "COVARIANCE.S",
  "DAYS",
  "DROP",
  "EXPAND",
  "FORECAST.LINEAR",
  "FORMULATEXT",
  "HSTACK",
  "IFNA",
  "IFS",
  "ISFORMULA",
  "MAXIFS",
  "MINIFS",
  "MODE.MULT",
  "MODE.SNGL",
  "PERCENTILE.EXC",
  "PERCENTILE.INC",
  "QUARTILE.EXC",
  "QUARTILE.INC",
  "RANK.AVG",
  "RANK.EQ",
  "SEQUENCE",
  "STDEV.P",
  "STDEV.S",
  "SWITCH",
  "TAKE",
  "TEXTAFTER",
  "TEXTBEFORE",
  "TEXTJOIN",
  "TEXTSPLIT",
  "TOCOL",
  "TOROW",
  "UNICHAR",
  "UNICODE",
  "UNIQUE",
  "VAR.P",
  "VAR.S",
  "VSTACK",
  "WRAPCOLS",
  "WRAPROWS",
  "XLOOKUP",
  "XMATCH",
  "XOR",
];

const XLWS_FUNCTIONS = ["FILTER", "SORT"];

const XLFN_FUNCTION_PATTERN = XLFN_FUNCTIONS
  .map((name) => name.replaceAll(".", "\\."))
  .join("|");

const MODEL_FUNCTION = new RegExp(`(?<![A-Za-z0-9_.])(${XLFN_FUNCTION_PATTERN})(?=\\s*\\()`, "giu");
const PACKAGE_FUNCTION = new RegExp(`(?<![A-Za-z0-9_.])_xlfn\\.(?:_xlws\\.)?(${XLFN_FUNCTION_PATTERN})(?=\\s*\\()`, "giu");
const XLWS_FUNCTION_PATTERN = XLWS_FUNCTIONS.join("|");
const MODEL_XLWS_FUNCTION = new RegExp(`(?<![A-Za-z0-9_.])(${XLWS_FUNCTION_PATTERN})(?=\\s*\\()`, "giu");
const PACKAGE_XLWS_FUNCTION = new RegExp(`(?<![A-Za-z0-9_.])_xlfn\\.(?:_xlws\\.)?(${XLWS_FUNCTION_PATTERN})(?=\\s*\\()`, "giu");

const SPILL_REFERENCE = /(?<![A-Za-z0-9_.])(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_. ]*))!)?(\$?[A-Za-z]{1,3}\$?[1-9]\d*)#(?![A-Za-z0-9_.])/gu;
const ANCHOR_ARRAY = /(?<![A-Za-z0-9_.])_xlfn\.ANCHORARRAY\(\s*((?:(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_. ]*)!)?\$?[A-Za-z]{1,3}\$?[1-9]\d*)\s*\)/giu;

function mapFormulaCode(formula, transform, { protectSingleQuotes = true } = {}) {
  const source = String(formula || "");
  let output = "";
  let index = 0;
  let codeStart = 0;
  while (index < source.length) {
    if (source[index] !== '"' && (!protectSingleQuotes || source[index] !== "'") && source[index] !== "[") {
      index += 1;
      continue;
    }
    output += transform(source.slice(codeStart, index));
    const literalStart = index;
    if (source[index] === '"' || source[index] === "'") {
      const quote = source[index];
      index += 1;
      while (index < source.length) {
        if (source[index] !== quote) {
          index += 1;
          continue;
        }
        if (source[index + 1] === quote) {
          index += 2;
          continue;
        }
        index += 1;
        break;
      }
    } else {
      let depth = 0;
      while (index < source.length) {
        if (source[index] === "'" && /['\[\]#@]/u.test(source[index + 1] || "")) {
          index += 2;
          continue;
        }
        if (source[index] === "[") depth += 1;
        if (source[index] === "]") {
          depth -= 1;
          index += 1;
          if (depth === 0) break;
          continue;
        }
        index += 1;
      }
    }
    output += source.slice(literalStart, index);
    codeStart = index;
  }
  return output + transform(source.slice(codeStart));
}

function isWorksheetAddress(address) {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/u.exec(address);
  if (!match || Number(match[2]) > 1_048_576) return false;
  const column = [...match[1].toUpperCase()].reduce((value, character) => (value * 26) + character.charCodeAt(0) - 64, 0);
  return column <= 16_384;
}

export function xlsxFormulaFromModel(formula) {
  const packageLet = xlsxLetFormulaFromModel(formula);
  const packageSpills = mapFormulaCode(packageLet, (code) => code.replace(SPILL_REFERENCE, (_match, quotedSheet, bareSheet, address) => {
      if (!isWorksheetAddress(address)) return _match;
      const sheet = quotedSheet != null ? `'${quotedSheet}'!` : bareSheet != null ? `${bareSheet}!` : "";
      return `_xlfn.ANCHORARRAY(${sheet}${address})`;
    }), { protectSingleQuotes: false });
  return mapFormulaCode(packageSpills, (code) => code
    .replace(MODEL_XLWS_FUNCTION, "_xlfn._xlws.$1")
    .replace(MODEL_FUNCTION, "_xlfn.$1"));
}

export function modelFormulaFromXlsx(formula) {
  const modelSpills = mapFormulaCode(formula, (code) => code.replace(ANCHOR_ARRAY, "$1#"), { protectSingleQuotes: false });
  const modelFunctions = mapFormulaCode(modelSpills, (code) => code
    .replace(PACKAGE_XLWS_FUNCTION, "$1")
    .replace(PACKAGE_FUNCTION, "$1"));
  return modelLetFormulaFromXlsx(modelFunctions);
}
