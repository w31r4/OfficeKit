const LET_MAX_BINDINGS = 16;
const LET_TRANSFORM_MAX_FORMULA_LENGTH = 8192;
const LET_TRANSFORM_MAX_NESTING = 64;
const LET_NAME = /^[A-Za-z_][A-Za-z0-9_.]*$/u;
const FORMULA_IDENTIFIER_START = /[A-Za-z_]/u;
const FORMULA_IDENTIFIER_PART = /[A-Za-z0-9_.]/u;
const FORMULA_ERROR = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|UNKNOWN!|CONNECT!|CYCLE!)/iu;

function formulaWithinBudget(text) {
  if (text.length > LET_TRANSFORM_MAX_FORMULA_LENGTH) return false;
  let depth = 0;
  let bracketDepth = 0;
  let inString = false;
  let inSheetName = false;
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (inString) {
      if (character === '"') {
        if (text[index + 1] === '"') index += 1;
        else inString = false;
      }
      continue;
    }
    if (inSheetName) {
      if (character === "'") {
        if (text[index + 1] === "'") index += 1;
        else inSheetName = false;
      }
      continue;
    }
    if (character === '"') { inString = true; continue; }
    if (character === "'") { inSheetName = true; continue; }
    if (character === "[") { bracketDepth += 1; continue; }
    if (character === "]") { bracketDepth -= 1; if (bracketDepth < 0) return false; continue; }
    if (bracketDepth > 0) continue;
    if (character === "(") { depth += 1; if (depth > LET_TRANSFORM_MAX_NESTING) return false; }
    else if (character === ")") { depth -= 1; if (depth < 0) return false; }
  }
  return depth === 0 && bracketDepth === 0 && !inString && !inSheetName;
}

function matchingParenthesis(text, openIndex) {
  let depth = 0;
  let bracketDepth = 0;
  let inString = false;
  let inSheetName = false;
  for (let index = openIndex; index < text.length; index += 1) {
    const character = text[index];
    if (inString) {
      if (character === '"') {
        if (text[index + 1] === '"') index += 1;
        else inString = false;
      }
      continue;
    }
    if (inSheetName) {
      if (character === "'") {
        if (text[index + 1] === "'") index += 1;
        else inSheetName = false;
      }
      continue;
    }
    if (character === '"') { inString = true; continue; }
    if (character === "'") { inSheetName = true; continue; }
    if (character === "[") { bracketDepth += 1; continue; }
    if (character === "]") { bracketDepth = Math.max(0, bracketDepth - 1); continue; }
    if (bracketDepth > 0) continue;
    if (character === "(") depth += 1;
    else if (character === ")") {
      depth -= 1;
      if (depth === 0) return index;
      if (depth < 0) return undefined;
    }
  }
  return undefined;
}

function splitFormulaArguments(text) {
  const arguments_ = [];
  let start = 0;
  let depth = 0;
  let bracketDepth = 0;
  let inString = false;
  let inSheetName = false;
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (inString) {
      if (character === '"') {
        if (text[index + 1] === '"') index += 1;
        else inString = false;
      }
      continue;
    }
    if (inSheetName) {
      if (character === "'") {
        if (text[index + 1] === "'") index += 1;
        else inSheetName = false;
      }
      continue;
    }
    if (character === '"') { inString = true; continue; }
    if (character === "'") { inSheetName = true; continue; }
    if (character === "[") { bracketDepth += 1; continue; }
    if (character === "]") { bracketDepth -= 1; if (bracketDepth < 0) return undefined; continue; }
    if (bracketDepth > 0) continue;
    if (character === "(") depth += 1;
    else if (character === ")") { depth -= 1; if (depth < 0) return undefined; }
    else if (character === "," && depth === 0) { arguments_.push(text.slice(start, index)); start = index + 1; }
  }
  if (depth !== 0 || bracketDepth !== 0 || inString || inSheetName) return undefined;
  arguments_.push(text.slice(start));
  return arguments_;
}

function bareLetName(value) {
  const trimmed = String(value).trim();
  const name = trimmed.replace(/^_xlpm\./iu, "");
  if (!LET_NAME.test(name) || /^(?:R|C)$/iu.test(name) || /^[A-Za-z]+[1-9]\d*$/u.test(name)) return undefined;
  return name;
}

function replaceTrimmed(value, replacement) {
  const leading = /^\s*/u.exec(value)?.[0] || "";
  const trailing = /\s*$/u.exec(value)?.[0] || "";
  return `${leading}${replacement}${trailing}`;
}

function nextNonWhitespace(text, index) {
  while (index < text.length && /\s/u.test(text[index])) index += 1;
  return { index, character: text[index] };
}

function previousNonWhitespace(text, index) {
  index -= 1;
  while (index >= 0 && /\s/u.test(text[index])) index -= 1;
  return text[index];
}

function isNumericExponent(text, start, identifier, next) {
  return /^E$/iu.test(identifier)
    && /[0-9.]/u.test(text[start - 1] || "")
    && (/[0-9]/u.test(next.character || "")
      || ((next.character === "+" || next.character === "-") && /[0-9]/u.test(text[next.index + 1] || "")));
}

function transformExpression(text, scope, direction, state, packageScope = false) {
  let output = "";
  for (let index = 0; index < text.length;) {
    const character = text[index];
    if (character === "#") {
      const error = FORMULA_ERROR.exec(text.slice(index))?.[0];
      if (error) { output += error; index += error.length; continue; }
    }
    if (character === '"' || character === "'") {
      const quote = character;
      const start = index;
      index += 1;
      while (index < text.length) {
        if (text[index] !== quote) { index += 1; continue; }
        if (text[index + 1] === quote) { index += 2; continue; }
        index += 1;
        break;
      }
      output += text.slice(start, index);
      continue;
    }
    if (character === "[") {
      const start = index;
      let depth = 0;
      while (index < text.length) {
        if (text[index] === "'" && /['\[\]#@]/u.test(text[index + 1] || "")) { index += 2; continue; }
        if (text[index] === "[") depth += 1;
        if (text[index] === "]") {
          depth -= 1;
          index += 1;
          if (depth === 0) break;
          continue;
        }
        index += 1;
      }
      output += text.slice(start, index);
      continue;
    }
    if (!FORMULA_IDENTIFIER_START.test(character)) {
      output += character;
      index += 1;
      continue;
    }

    const start = index;
    index += 1;
    while (index < text.length && FORMULA_IDENTIFIER_PART.test(text[index])) index += 1;
    const identifier = text.slice(start, index);
    const next = nextNonWhitespace(text, index);
    if (next.character === "(") {
      const closeIndex = matchingParenthesis(text, next.index);
      if (closeIndex == null) { state.valid = false; return text; }
      const callText = text.slice(start, closeIndex + 1);
      if (/^(?:_xlfn\.)?LET$/iu.test(identifier)) output += transformLetCall(callText, scope, direction, state);
      else {
        output += identifier;
        output += text.slice(index, next.index + 1);
        output += transformExpression(text.slice(next.index + 1, closeIndex), scope, direction, state, packageScope);
        output += ")";
      }
      index = closeIndex + 1;
      continue;
    }

    const packageName = /^_xlpm\.(.+)$/iu.exec(identifier)?.[1];
    const logicalName = (packageName || identifier).toUpperCase();
    const isLocalReference = scope.has(logicalName)
      && previousNonWhitespace(text, start) !== "!"
      && previousNonWhitespace(text, start) !== "#"
      && next.character !== "!"
      && next.character !== "["
      && !isNumericExponent(text, start, identifier, next);
    if (direction === "model" && packageScope && packageName != null && !isLocalReference) {
      state.valid = false;
      return text;
    }
    if (direction === "model" && packageScope && packageName == null && isLocalReference) {
      state.valid = false;
      return text;
    }
    if (direction === "package" && isLocalReference && packageName == null) output += `_xlpm.${identifier}`;
    else if (direction === "model" && isLocalReference && packageName != null) output += packageName;
    else output += identifier;
  }
  return output;
}

function transformLetCall(callText, inheritedScope, direction, state) {
  const match = /^((?:_xlfn\.)?LET)(\s*)\(/iu.exec(callText);
  if (!match) return callText;
  const openIndex = match[0].length - 1;
  const closeIndex = matchingParenthesis(callText, openIndex);
  if (closeIndex !== callText.length - 1) { state.valid = false; return callText; }
  const arguments_ = splitFormulaArguments(callText.slice(openIndex + 1, closeIndex));
  if (!arguments_ || arguments_.length < 3 || arguments_.length % 2 === 0 || (arguments_.length - 1) / 2 > LET_MAX_BINDINGS) {
    state.valid = false;
    return callText;
  }
  const packageScope = direction === "model"
    && (/^_xlfn\.LET$/iu.test(match[1]) || arguments_.slice(0, -1).some((argument, index) => index % 2 === 0 && /^\s*_xlpm\./iu.test(argument)));
  if (direction === "model" && state.letMode && state.letMode !== (packageScope ? "package" : "model")) {
    state.valid = false;
    return callText;
  }
  if (direction === "model") state.letMode = packageScope ? "package" : "model";
  const names = [];
  for (let index = 0; index < arguments_.length - 1; index += 2) {
    if (packageScope && !/^\s*_xlpm\./iu.test(arguments_[index])) { state.valid = false; return callText; }
    const name = bareLetName(arguments_[index]);
    if (!name) { state.valid = false; return callText; }
    names.push(name);
  }

  const scope = new Set(inheritedScope);
  const transformed = [];
  for (let index = 0; index < arguments_.length - 1; index += 2) {
    const name = names[index / 2];
    transformed.push(replaceTrimmed(arguments_[index], direction === "package" ? `_xlpm.${name}` : name));
    transformed.push(transformExpression(arguments_[index + 1], scope, direction, state, packageScope));
    scope.add(name.toUpperCase());
  }
  transformed.push(transformExpression(arguments_.at(-1), scope, direction, state, packageScope));
  const functionName = direction === "package" ? "_xlfn.LET" : "LET";
  return `${functionName}${match[2]}(${transformed.join(",")})`;
}

function transformLetFormula(formula, direction) {
  const source = String(formula || "");
  if (!formulaWithinBudget(source)) return source;
  const state = { valid: true, letMode: undefined };
  const transformed = transformExpression(source, new Set(), direction, state);
  return state.valid ? transformed : source;
}

export function xlsxLetFormulaFromModel(formula) {
  return transformLetFormula(formula, "package");
}

export function modelLetFormulaFromXlsx(formula) {
  return transformLetFormula(formula, "model");
}
