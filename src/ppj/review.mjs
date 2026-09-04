import { createHash } from "node:crypto";

import { canonicalJson, normalizePresentationAuthoringPlan } from "../cli/authoring-plan.mjs";
import { projectPptxToPpj } from "./native.mjs";

function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

function issue(type, message, severity = "error", details = {}) {
  return { kind: "reviewIssue", type, severity, message, ...details };
}

function hard(issues) {
  return issues.some((entry) => entry.severity === "error");
}

function section(issues, extra = {}) {
  const ok = !hard(issues);
  return {
    status: ok ? issues.length ? "passed-with-warnings" : "passed" : "failed",
    ok,
    issues,
    ...extra,
  };
}

function boundedText(value, maximum) {
  const text = String(value || "");
  if (text.length <= maximum) return { text, truncated: false };
  return { text: `${text.slice(0, Math.max(0, maximum - 32))}\n{\"kind\":\"truncated\"}`, truncated: true };
}

function parseProgram(receipt) {
  const source = Buffer.from(receipt.programJson || []).toString("utf8");
  const program = JSON.parse(source);
  if (!program || typeof program !== "object" || !Array.isArray(program.pages)) {
    throw new Error("Native PPJ projection returned an invalid canonical program.");
  }
  return program;
}

function walkElements(elements, page, records, parentId = null) {
  if (!Array.isArray(elements)) return;
  for (const element of elements) {
    if (!element || typeof element !== "object") continue;
    records.push({ page, parentId, element });
    const childId = typeof element.id === "string" ? element.id : parentId;
    walkElements(element.elements, page, records, childId);
    walkElements(element.children, page, records, childId);
    walkElements(element.items, page, records, childId);
  }
}

function programRecords(program) {
  const records = [];
  for (const [index, page] of program.pages.entries()) {
    walkElements(page.elements, index + 1, records);
  }
  return records;
}

function recordCounts(program, records, receipt) {
  const counts = { page: program.pages.length };
  for (const { element } of records) {
    const kind = `element:${String(element.type || element.kind || "unknown")}`;
    counts[kind] = (counts[kind] || 0) + 1;
  }
  if (Array.isArray(program.assets)) counts.asset = program.assets.length;
  if (receipt.expandedElementCount > records.length) counts.expandedElement = receipt.expandedElementCount;
  return counts;
}

function semanticInspection(program, records, receipt, maxChars) {
  const lines = [
    { kind: "summary", schema: program.schema, pages: program.pages.length, sourceBound: receipt.sourceBound, expandedElementCount: receipt.expandedElementCount },
    ...program.pages.map((page, index) => ({ kind: "page", index: index + 1, id: page.id, name: page.name, elements: records.filter((record) => record.page === index + 1).length })),
  ].map((entry) => JSON.stringify(entry)).join("\n");
  return boundedText(lines, maxChars);
}

function nativeIssues(receipt) {
  return (receipt.diagnostics || []).map((diagnostic) => issue(
    diagnostic.code || "ppjDiagnostic",
    diagnostic.message || "Native PPJ diagnostic.",
    Number(diagnostic.severity) >= 3 ? "error" : Number(diagnostic.severity) === 2 ? "warning" : "info",
    { sourcePath: diagnostic.sourcePath || undefined, sourceIdentity: diagnostic.sourceIdentity || undefined },
  ));
}

function frameValues(element) {
  const frame = element?.frame;
  if (!frame || typeof frame !== "object") return null;
  const values = [frame.x, frame.y, frame.width, frame.height].map(Number);
  return values.every(Number.isFinite) ? values : null;
}

function normalizedDegrees(value) {
  const degrees = Number(value);
  if (!Number.isFinite(degrees)) return 0;
  const normalized = ((degrees % 360) + 360) % 360;
  return normalized > 180 ? normalized - 360 : normalized;
}

function rotatedBounds(frame, rotation) {
  const [x, y, width, height] = frame;
  const radians = rotation * Math.PI / 180;
  if (Math.abs(radians) < 1e-12) return [...frame];
  const centerX = x + width / 2;
  const centerY = y + height / 2;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const corners = [[-width / 2, -height / 2], [width / 2, -height / 2], [width / 2, height / 2], [-width / 2, height / 2]];
  const points = corners.map(([localX, localY]) => [
    centerX + localX * cos - localY * sin,
    centerY + localX * sin + localY * cos,
  ]);
  const minX = Math.min(...points.map(([pointX]) => pointX));
  const minY = Math.min(...points.map(([, pointY]) => pointY));
  const maxX = Math.max(...points.map(([pointX]) => pointX));
  const maxY = Math.max(...points.map(([, pointY]) => pointY));
  return [minX, minY, maxX - minX, maxY - minY];
}

function shadowValues(shadow) {
  if (!shadow || typeof shadow !== "object") return null;
  const blur = Number(shadow.blur);
  const distance = Number(shadow.distance);
  const angle = Number(shadow.angle);
  if (![blur, distance, angle].every(Number.isFinite) || blur < 0 || distance < 0) return null;
  return { blur, distance, angle: normalizedDegrees(angle), rotateWithShape: shadow.rotateWithShape === true };
}

function shadowCandidates(element) {
  const candidates = [
    element?.shadow,
    element?.style?.shadow,
    element?.style?.frame?.shadow,
    element?.style?.titleTextStyle?.shadow,
    element?.style?.legendTextStyle?.shadow,
    element?.style?.dataLabels?.textStyle?.shadow,
  ];
  return [...new Set(candidates.filter(Boolean))];
}

function unionBounds(left, right) {
  const minX = Math.min(left[0], right[0]);
  const minY = Math.min(left[1], right[1]);
  const maxX = Math.max(left[0] + left[2], right[0] + right[2]);
  const maxY = Math.max(left[1] + left[3], right[1] + right[3]);
  return [minX, minY, maxX - minX, maxY - minY];
}

function visualBounds(element) {
  const frame = frameValues(element);
  if (!frame) return null;
  const rotation = normalizedDegrees(element?.frame?.rotation);
  let bbox = rotatedBounds(frame, rotation);
  let hasRotation = Math.abs(rotation) > 1e-9;
  let hasShadow = false;
  for (const candidate of shadowCandidates(element)) {
    const shadow = shadowValues(candidate);
    if (!shadow) continue;
    const angle = (shadow.angle + (shadow.rotateWithShape ? rotation : 0)) * Math.PI / 180;
    const offsetX = shadow.distance * Math.cos(angle);
    const offsetY = shadow.distance * Math.sin(angle);
    const blur = shadow.blur;
    const shadowBox = [bbox[0] + offsetX - blur, bbox[1] + offsetY - blur, bbox[2] + blur * 2, bbox[3] + blur * 2];
    bbox = unionBounds(bbox, shadowBox);
    hasShadow = true;
  }
  return { bbox, frame, rotation, hasRotation, hasShadow };
}

function hasVisibleText(element) {
  const text = element?.text;
  if (typeof text === "string") return text.length > 0;
  if (!text || typeof text !== "object") return false;
  if (typeof text.plainText === "string" && text.plainText.length > 0) return true;
  return Array.isArray(text.paragraphs) && text.paragraphs.some((paragraph) =>
    Array.isArray(paragraph?.runs) && paragraph.runs.some((run) => typeof run?.text === "string" && run.text.length > 0));
}

function textLines(element) {
  const text = element?.text;
  if (typeof text === "string") return text.split("\n");
  if (!text || typeof text !== "object") return [];
  if (typeof text.plainText === "string") return text.plainText.split("\n");
  if (!Array.isArray(text.paragraphs)) return [];
  return text.paragraphs.map((paragraph) => Array.isArray(paragraph?.runs)
    ? paragraph.runs.map((run) => typeof run?.text === "string" ? run.text : "").join("")
    : "");
}

function firstFinite(...values) {
  return values.map(Number).find((value) => Number.isFinite(value) && value > 0) ?? null;
}

function textFontSize(element) {
  const sizes = [];
  const collect = (value) => {
    if (!value || typeof value !== "object") return;
    if (Array.isArray(value)) {
      for (const item of value) collect(item);
      return;
    }
    for (const [key, child] of Object.entries(value)) {
      if (["size", "fontSize", "fontSizePoints"].includes(key) && Number.isFinite(Number(child))) sizes.push(Number(child));
      else if (child && typeof child === "object") collect(child);
    }
  };
  collect(element?.textStyle);
  collect(element?.style);
  collect(element?.text);
  return firstFinite(...sizes) ?? 18;
}

function textMargins(element) {
  const margins = element?.textStyle?.margins || element?.style?.margins;
  if (!margins || typeof margins !== "object") return { left: 0, right: 0, top: 0, bottom: 0 };
  return {
    left: Math.max(0, Number(margins.left) || 0),
    right: Math.max(0, Number(margins.right) || 0),
    top: Math.max(0, Number(margins.top) || 0),
    bottom: Math.max(0, Number(margins.bottom) || 0),
  };
}

function estimatedTextWidth(value, fontSize) {
  let width = 0;
  for (const character of String(value)) {
    if (/\s/u.test(character)) width += fontSize * 0.32;
    else if (/[\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff]/u.test(character)) width += fontSize;
    else width += fontSize * 0.56;
  }
  return width;
}

function textMeasurement(element) {
  if (!hasVisibleText(element)) return null;
  const frame = frameValues(element);
  if (!frame) return null;
  const lines = textLines(element);
  if (!lines.length) return null;
  const fontSize = textFontSize(element);
  const margins = textMargins(element);
  const availableWidth = Math.max(0, frame[2] - margins.left - margins.right);
  const availableHeight = Math.max(0, frame[3] - margins.top - margins.bottom);
  const lineWidths = lines.map((line) => estimatedTextWidth(line, fontSize));
  const widestLine = Math.max(...lineWidths, 0);
  const wrappedLines = availableWidth > 0
    ? lineWidths.reduce((sum, width) => sum + Math.max(1, Math.ceil(width / availableWidth)), 0)
    : lines.length;
  const lineHeight = fontSize * 1.2;
  const measuredWidth = widestLine;
  const measuredHeight = wrappedLines * lineHeight;
  const overflowWidth = widestLine > availableWidth + 0.01;
  const overflowHeight = measuredHeight > availableHeight + 0.01;
  const autoFit = element?.textStyle?.autoFit || element?.style?.autoFit || "none";
  return {
    method: "deterministic-character-metric",
    fontSize,
    lineCount: lines.length,
    wrappedLines,
    widestLine,
    measuredWidth,
    measuredHeight,
    availableWidth,
    availableHeight,
    margins,
    autoFit,
    overflowWidth,
    overflowHeight,
    overflow: overflowWidth || overflowHeight,
  };
}

function isDecorative(element) {
  return element?.accessibility?.decorative === true && !hasVisibleText(element);
}

function pathValue(value, path) {
  let cursor = value;
  for (const part of String(path || "").split(".").filter(Boolean)) {
    if (!cursor || typeof cursor !== "object" || !Object.hasOwn(cursor, part)) return { found: false, value: undefined };
    cursor = cursor[part];
  }
  return { found: true, value: cursor };
}

function grammarStyles(program, element) {
  const styles = program.design?.styles;
  if (!styles || typeof styles !== "object") return [];
  const family = element?.type === "text" || element?.type === "placeholder" ? "text"
    : element?.type === "chart" ? "chart"
      : element?.type === "table" ? "table" : "shape";
  const styleRef = element?.styleRef;
  if (typeof styleRef !== "string") return [];
  const entries = Array.isArray(styles[family]) ? styles[family] : [];
  const match = entries.find((entry) => entry?.id === styleRef && entry?.style && typeof entry.style === "object");
  return match ? [match.style] : [];
}

function grammarThemeValue(theme, target) {
  const direct = pathValue(theme, target);
  if (direct.found) return direct;
  const leaf = String(target || "").split(".").pop();
  const colors = Array.isArray(theme?.colors) ? theme.colors : [];
  const color = colors.find((entry) => entry?.id === leaf);
  return color ? { found: true, value: color.value } : { found: false, value: undefined };
}

function grammarSourceValue(program, record, target, source, tokenMap) {
  const resolveToken = (result) => {
    if (!result?.found || !result.value || typeof result.value !== "object" || Array.isArray(result.value)) return result;
    const keys = Object.keys(result.value);
    if (keys.length !== 1 || keys[0] !== "token" || typeof result.value.token !== "string") return result;
    const token = tokenMap.get(result.value.token);
    return token
      ? { found: true, value: token.value, token: result.value.token, kind: token.kind }
      : result;
  };
  const element = record?.element;
  if (source === "inline") {
    for (const owner of [element, element?.style]) {
      const result = pathValue(owner, target);
      if (result.found) return resolveToken(result);
    }
    return { found: false, value: undefined };
  }
  if (source === "styleRef") {
    for (const style of grammarStyles(program, element)) {
      const result = pathValue(style, target);
      if (result.found) return resolveToken(result);
    }
    return { found: false, value: undefined };
  }
  if (source === "theme") return resolveToken(grammarThemeValue(program.design?.theme, target));
  if (source === "master") {
    const page = Number(record?.page) > 0 ? program.pages?.[Number(record.page) - 1] : undefined;
    const layoutId = page?.layout;
    const layout = Array.isArray(program.design?.layouts) ? program.design.layouts.find((entry) => entry?.id === layoutId) : undefined;
    const master = layout && Array.isArray(program.design?.masters) ? program.design.masters.find((entry) => entry?.id === layout.master) : undefined;
    for (const owner of [layout, master]) {
      const result = pathValue(owner, target);
      if (result.found) return resolveToken(result);
    }
    return { found: false, value: undefined };
  }
  if (source === "default") {
    const exact = tokenMap.get(target);
    if (exact) return { found: true, value: exact.value, token: target, kind: exact.kind };
    const leaf = String(target || "").split(".").pop();
    const fallback = tokenMap.get(leaf);
    if (fallback) return { found: true, value: fallback.value, token: leaf, kind: fallback.kind };
  }
  return { found: false, value: undefined };
}

function grammarEqual(left, right) {
  return canonicalJson(left) === canonicalJson(right);
}

function predicateMatches(actual, predicate) {
  const op = predicate?.op;
  const expected = predicate?.value;
  if (op === "in") return Array.isArray(expected) && expected.some((value) => grammarEqual(actual, value));
  if (op === "eq") return grammarEqual(actual, expected);
  if (op === "neq") return !grammarEqual(actual, expected);
  const left = Number(actual);
  const right = Number(expected);
  if (!Number.isFinite(left) || !Number.isFinite(right)) return false;
  return op === "gt" ? left > right : op === "gte" ? left >= right : op === "lt" ? left < right : op === "lte" ? left <= right : false;
}

function grammarReview(program, records) {
  const grammar = program.design?.grammar;
  if (!grammar || typeof grammar !== "object") return { status: "not-declared", ok: true, scope: "typed design grammar evaluation (read-only)", issues: [], tokens: {}, resolutions: [], predicates: [] };
  const issues = [];
  const tokenMap = new Map();
  const tokenOutput = {};
  if (grammar.tokens && typeof grammar.tokens === "object" && !Array.isArray(grammar.tokens)) {
    for (const [id, definition] of Object.entries(grammar.tokens)) {
      if (!definition || typeof definition !== "object") continue;
      const entry = { kind: definition.kind, value: definition.value };
      tokenMap.set(id, entry);
      tokenOutput[id] = entry;
    }
  }
  const resolutions = [];
  const precedence = Array.isArray(grammar.stylePrecedence) ? grammar.stylePrecedence : [];
  for (const rule of precedence) {
    if (!rule || typeof rule !== "object" || typeof rule.target !== "string" || !Array.isArray(rule.sources)) continue;
    const entries = [];
    for (const record of records) {
      const selected = rule.sources.map((source) => ({ source, result: grammarSourceValue(program, record, rule.target, source, tokenMap) }))
        .find((candidate) => candidate.result.found);
      if (!selected) continue;
      entries.push({ page: record.page, id: record.element?.id, source: selected.source, value: selected.result.value, ...(selected.result.token ? { token: selected.result.token, kind: selected.result.kind } : {}) });
    }
    resolutions.push({ target: rule.target, sources: rule.sources, entries });
  }
  const predicates = [];
  const declaredPredicates = Array.isArray(grammar.predicates) ? grammar.predicates : [];
  for (const predicate of declaredPredicates) {
    if (!predicate || typeof predicate !== "object") continue;
    const checks = [];
    for (const record of records) {
      const field = String(predicate.field || "");
      const path = field.startsWith("element.") ? field.slice("element.".length) : field;
      const actual = pathValue(record.element, path);
      const pass = actual.found && predicateMatches(actual.value, predicate);
      checks.push({ page: record.page, id: record.element?.id, pass, ...(actual.found ? { actual: actual.value } : { missing: true }) });
      if (!pass) issues.push(issue("grammarPredicateViolation", `Grammar predicate ${predicate.id || "<anonymous>"} does not hold for ${record.element?.id || "<anonymous>"}.`, "warning", { predicateId: predicate.id, slide: record.page, id: record.element?.id, field: predicate.field }));
    }
    predicates.push({ id: predicate.id, field: predicate.field, op: predicate.op, checks });
  }
  return section(issues, {
    scope: "typed design grammar evaluation (read-only)",
    precedenceOrder: "first declared source wins",
    tokens: tokenOutput,
    resolutions,
    predicates,
  });
}

function contains(parentId, candidateId, recordsById) {
  let cursor = candidateId;
  while (recordsById.has(cursor)) {
    const parent = recordsById.get(cursor).parentId;
    if (!parent) return false;
    if (parent === parentId) return true;
    cursor = parent;
  }
  return false;
}

function overlapArea(left, right) {
  const x = Math.max(0, Math.min(left[0] + left[2], right[0] + right[2]) - Math.max(left[0], right[0]));
  const y = Math.max(0, Math.min(left[1] + left[3], right[1] + right[3]) - Math.max(left[1], right[1]));
  return x * y;
}

function layoutReview(program, records, options = {}) {
  const issues = [];
  const width = Number(program.design?.canvas?.width ?? program.design?.width);
  const height = Number(program.design?.canvas?.height ?? program.design?.height);
  const minOverlapArea = Number.isFinite(Number(options.minOverlapArea)) ? Math.max(0, Number(options.minOverlapArea)) : 64;
  const boundsPadding = Number.isFinite(Number(options.boundsPadding)) ? Math.max(0, Number(options.boundsPadding)) : 0;
  for (const { page, element } of records) {
    const frame = element.frame;
    if (!frame || typeof frame !== "object") continue;
    const values = [frame.x, frame.y, frame.width, frame.height].map(Number);
    if (!values.every(Number.isFinite)) {
      issues.push(issue("invalidFrame", `Element ${element.id || "<anonymous>"} on page ${page} has a non-finite frame.`, "error", { slide: page, id: element.id }));
      continue;
    }
    const [x, y, w, h] = values;
    if (w < 0 || h < 0) issues.push(issue("invalidFrameSize", `Element ${element.id || "<anonymous>"} on page ${page} has a negative frame size.`, "error", { slide: page, id: element.id }));
    const visual = visualBounds(element);
    const visualBox = visual?.bbox || [x, y, w, h];
    const text = textMeasurement(element);
    if (text?.overflow) {
      const autoFitHandled = ["shrink", "fit", "resizeShape"].includes(text.autoFit);
      issues.push(issue(
        "textOverflowEstimated",
        `Element ${element.id || "<anonymous>"} on page ${page} may overflow its text frame under the bounded text metric.`,
        autoFitHandled ? "info" : "warning",
        {
          slide: page,
          id: element.id,
          measurement: text,
          recommendation: autoFitHandled ? "Confirm the host AutoFit result." : "Increase the frame or provide an explicit AutoFit policy.",
        },
      ));
    }
    if (Number.isFinite(width) && Number.isFinite(height) && (visualBox[0] < -boundsPadding || visualBox[1] < -boundsPadding || visualBox[0] + visualBox[2] > width + boundsPadding || visualBox[1] + visualBox[3] > height + boundsPadding)) {
      issues.push(issue("frameOutsideCanvas", `Element ${element.id || "<anonymous>"} on page ${page} extends beyond the PPJ canvas.`, "warning", { slide: page, id: element.id, bbox: visualBox, frame: [x, y, w, h], visualBounds: { rotation: visual.rotation, shadow: visual.hasShadow } }));
    }
  }
  const canvasArea = Number.isFinite(width) && Number.isFinite(height) ? Math.max(0, width * height) : 0;
  for (const page of new Set(records.map((record) => record.page))) {
    const pageRecords = records
      .filter((record) => record.page === page && record.element?.id && record.element.type !== "connector" && !isDecorative(record.element))
      .map((record, order) => ({ ...record, order, visual: visualBounds(record.element) }))
      .map((record) => ({ ...record, values: record.visual?.bbox || null }))
      .filter((record) => record.values && record.values[2] > 0 && record.values[3] > 0);
    const recordsById = new Map(pageRecords.map((record) => [record.element.id, record]));
    for (let leftIndex = 0; leftIndex < pageRecords.length; leftIndex += 1) {
      const left = pageRecords[leftIndex];
      const leftArea = left.values[2] * left.values[3];
      if (canvasArea > 0 && leftArea / canvasArea >= 0.8) continue;
      for (let rightIndex = leftIndex + 1; rightIndex < pageRecords.length; rightIndex += 1) {
        const right = pageRecords[rightIndex];
        if (contains(left.element.id, right.element.id, recordsById) || contains(right.element.id, left.element.id, recordsById)) continue;
        const rightArea = right.values[2] * right.values[3];
        if (canvasArea > 0 && rightArea / canvasArea >= 0.8) continue;
        const area = overlapArea(left.values, right.values);
        if (area < minOverlapArea) continue;
        issues.push(issue(
          "elementOverlap",
          `Elements ${left.element.id} and ${right.element.id} on page ${page} overlap by about ${Math.round(area)}px².`,
          "warning",
          {
            slide: page,
            ids: [left.element.id, right.element.id],
            zOrder: [left.order, right.order],
            overlapArea: Math.round(area),
            bbox: [left.values, right.values],
            detection: left.visual.hasRotation || left.visual.hasShadow || right.visual.hasRotation || right.visual.hasShadow
              ? "rotated-visual-bounds"
              : "axis-aligned-frame",
            frames: [left.visual.frame, right.visual.frame],
            visualBounds: [left.values, right.values],
          },
        ));
      }
    }
  }
  return section(issues, { scope: "PPJ visual bounds, z-order, and overlap (read-only review)" });
}

function accessibilityReview(program, records) {
  const issues = [];
  const pages = [];
  const directIdsForPage = (page) => (Array.isArray(page?.elements) ? page.elements : [])
    .map((element) => element && typeof element === "object" ? element.id : undefined)
    .filter((id) => typeof id === "string");
  const alternativeText = (element) => {
    const accessibility = element?.accessibility;
    if (!accessibility || typeof accessibility !== "object") return false;
    return [accessibility.title, accessibility.description]
      .some((value) => typeof value === "string" && value.trim().length > 0);
  };
  const semanticAltKinds = new Set(["image", "chart", "table", "smartArt", "media", "ole"]);

  for (const [index, page] of program.pages.entries()) {
    const directIds = directIdsForPage(page);
    const declared = Array.isArray(page?.readingOrder) ? page.readingOrder.filter((id) => typeof id === "string") : null;
    const explicit = declared !== null;
    const effective = declared ?? directIds;
    const complete = effective.length === directIds.length &&
      new Set(effective).size === effective.length &&
      new Set(effective).size === new Set(directIds).size &&
      effective.every((id) => directIds.includes(id));
    if (explicit && !complete) {
      issues.push(issue(
        "invalidReadingOrder",
        `Page ${page?.id || index + 1} readingOrder must be a complete permutation of direct element IDs.`,
        "error",
        { slide: index + 1, pageId: page?.id, readingOrder: declared, directElementIds: directIds },
      ));
    }
    pages.push({
      slide: index + 1,
      pageId: page?.id,
      readingOrder: complete ? effective : directIds,
      source: explicit && complete ? "explicit" : "implicit-native-order",
      complete,
    });
  }

  for (const { page, parentId, element } of records) {
    if (!semanticAltKinds.has(element?.type)) continue;
    const accessibility = element?.accessibility;
    if (accessibility?.decorative === true) {
      if (hasVisibleText(element)) {
        issues.push(issue(
          "decorativeElementHasText",
          `Element ${element.id || "<anonymous>"} is marked decorative but contains visible text.`,
          "warning",
          { slide: page, id: element.id, parentId },
        ));
      }
      continue;
    }
    if (!alternativeText(element)) {
      issues.push(issue(
        "missingAlternativeText",
        `Element ${element.id || "<anonymous>"} requires an accessibility title or description, or an explicit decorative=true decision.`,
        "warning",
        { slide: page, id: element.id, parentId, elementType: element.type },
      ));
    }
  }

  return section(issues, {
    scope: "PPJ accessibility metadata and reading order (machine review)",
    conformanceClaimed: false,
    pages,
    manualChecks: [
      { type: "group-and-smartArt-reading-order", status: "manual-review-required" },
      { type: "host-accessibility-checker", status: "not-run" },
      { type: "author-intent-and-decorative-meaning", status: "manual-review-required" },
    ],
  });
}

function pageSignature(page) {
  return sha256(Buffer.from(canonicalJson(page)));
}

function collectDesignTokens(value, state, key = "") {
  if (Array.isArray(value)) {
    for (const item of value) collectDesignTokens(item, state, key);
    return;
  }
  if (!value || typeof value !== "object") {
    if (typeof value !== "string") return;
    if (/color|fill|line|stroke/iu.test(key) && (/^#[0-9a-f]{6}$/iu.test(value) || /^[a-z][a-z0-9_-]*$/iu.test(value))) state.colors.add(value);
    if (/font|typeface/iu.test(key)) state.fonts.add(value);
    return;
  }
  for (const [childKey, child] of Object.entries(value)) collectDesignTokens(child, state, childKey);
}

function textStats(value, state, inText = false) {
  if (Array.isArray(value)) {
    for (const item of value) textStats(item, state, inText);
    return;
  }
  if (!value || typeof value !== "object") return;
  for (const [key, child] of Object.entries(value)) {
    const textContext = inText || key === "text" || key === "paragraphs" || key === "runs" || key === "textStyle";
    if (textContext && key === "text" && typeof child === "string") state.characters += child.length;
    if (textContext && ["size", "fontSize"].includes(key) && Number.isFinite(Number(child))) state.fontSizes.push(Number(child));
    textStats(child, state, textContext);
  }
}

function normalizeChangedPageIds(value, pages, issues) {
  if (value == null) return [];
  if (!Array.isArray(value)) {
    issues.push(issue("invalidChangedPageIds", "changedPageIds must be an array."));
    return [];
  }
  const known = new Set(pages.map((page) => page.id));
  const result = [...new Set(value.map(String))];
  for (const id of result) if (!known.has(id)) issues.push(issue("unknownChangedPage", `changedPageIds contains unknown page ${id}.`, "error", { pageId: id }));
  return result;
}

function designReview(program, records, options = {}) {
  const grammar = grammarReview(program, records);
  if (options.authoringPlan == null && options.changedPageIds == null) {
    return { status: "not-requested", ok: true, planSha256: null, changedPageIds: [], issues: [], grammar };
  }
  const issues = [];
  let normalized;
  try {
    normalized = normalizePresentationAuthoringPlan(options.authoringPlan);
  } catch (error) {
    return section([issue("invalidAuthoringPlan", error?.message || String(error))], { planSha256: null, changedPageIds: [] });
  }
  const plan = normalized.plan;
  if (plan.pages.length !== program.pages.length) {
    issues.push(issue("authoringPlanPageCount", `Authoring plan declares ${plan.pages.length} pages but the PPJ contains ${program.pages.length}.`, "error", { expected: plan.pages.length, actual: program.pages.length }));
  }
  for (const [index, unresolved] of plan.unresolved.entries()) {
    if (typeof unresolved === "string" || unresolved?.required !== false && unresolved?.blocking !== false) {
      issues.push(issue("requiredAuthoringDecision", `Authoring plan still contains required unresolved item ${index + 1}.`, "error", { unresolvedIndex: index }));
    }
  }
  const changedPageIds = normalizeChangedPageIds(options.changedPageIds, plan.pages, issues);
  const pageSignatures = program.pages.map((page, index) => ({ pageId: plan.pages[index]?.id || page.id || `page-${index + 1}`, sha256: pageSignature(page) }));
  for (const [index, pagePlan] of plan.pages.entries()) {
    const pageRecords = records.filter((record) => record.page === index + 1);
    const stats = { characters: 0, fontSizes: [] };
    textStats(program.pages[index], stats);
    const budget = pagePlan.contentBudget || {};
    if (budget.maxObjects != null && pageRecords.length > budget.maxObjects) issues.push(issue("contentBudgetObjects", `Page ${pagePlan.id} contains ${pageRecords.length} objects; budget is ${budget.maxObjects}.`, "error", { pageId: pagePlan.id, slide: index + 1 }));
    if (budget.maxCharacters != null && stats.characters > budget.maxCharacters) issues.push(issue("contentBudgetCharacters", `Page ${pagePlan.id} contains ${stats.characters} characters; budget is ${budget.maxCharacters}.`, "error", { pageId: pagePlan.id, slide: index + 1 }));
    const bodyFloor = Number(plan.design?.designGrammar?.typography?.minimumBodyFontSize);
    if (Number.isFinite(bodyFloor) && stats.fontSizes.some((size) => size < bodyFloor)) issues.push(issue("minimumFontSize", `Page ${pagePlan.id} contains text below the ${bodyFloor}pt body floor.`, "warning", { pageId: pagePlan.id, slide: index + 1 }));
  }
  if (changedPageIds.length && Array.isArray(options.baselineDesign?.pageSignatures)) {
    const changed = new Set(changedPageIds);
    const baseline = new Map(options.baselineDesign.pageSignatures.map((entry) => [entry.pageId, entry.sha256]));
    for (const entry of pageSignatures) if (!changed.has(entry.pageId) && baseline.get(entry.pageId) !== entry.sha256) {
      issues.push(issue("undeclaredPageChange", `Page ${entry.pageId} changed outside changedPageIds.`, "error", { pageId: entry.pageId }));
    }
  } else if (changedPageIds.length) {
    issues.push(issue("changedPageScopeUnverified", "No baseline PPJ page signatures were available to prove non-target page stability.", "warning"));
  }
  const tokens = { colors: new Set(), fonts: new Set() };
  collectDesignTokens(program.design, tokens);
  for (const page of program.pages) collectDesignTokens(page, tokens);
  const profile = { paletteDirect: [...tokens.colors].sort(), fonts: [...tokens.fonts].sort() };
  return section(issues, {
    planSha256: normalized.sha256,
    changedPageIds,
    pageSignatures,
    profile,
    grammar,
    strategy: { status: normalized.strategyStatus, deliveryMode: plan.brief.deliveryMode ?? "hybrid" },
  });
}

function motionReview(program, options = {}) {
  const playbackEvidence = options.playbackEvidence ?? "structural";
  if (!new Set(["structural", "keynote", "powerpoint"]).has(playbackEvidence)) throw new TypeError("playbackEvidence must be structural, keynote, or powerpoint.");
  const issues = [];
  const motionUnits = [];
  const morphPairs = [];
  let normalized;
  if (options.authoringPlan != null) {
    try { normalized = normalizePresentationAuthoringPlan(options.authoringPlan); }
    catch (error) { issues.push(issue("invalidMotionAuthoringPlan", error?.message || String(error))); }
  }
  let animationCount = 0;
  for (const [index, page] of program.pages.entries()) {
    const elements = [];
    walkElements(page.elements, index + 1, elements);
    const ids = new Set(elements.map(({ element }) => element.id).filter(Boolean));
    const animations = Array.isArray(page.animations) ? page.animations : [];
    animationCount += animations.length;
    for (const [order, animation] of animations.entries()) {
      const targetId = animation.target ?? animation.targetId;
      motionUnits.push({ slide: index + 1, pageId: page.id, order: order + 1, id: animation.id, targetId, effect: animation.effect, start: animation.start });
      if (targetId && !ids.has(targetId)) issues.push(issue("invalidMotionTarget", `Animation ${animation.id || order + 1} on page ${page.id || index + 1} targets missing element ${targetId}.`, "error", { slide: index + 1, id: animation.id }));
    }
    if (animations.length > 32) issues.push(issue("motionNodeLimit", `Page ${page.id || index + 1} exceeds the 32-unit PPJ motion limit.`, "error", { slide: index + 1 }));
    const planPage = normalized?.plan.pages[index];
    if (planPage?.motionIntent) {
      if (planPage.motionIntent.units.length !== animations.length) issues.push(issue("motionPlanMismatch", `Page ${planPage.id} declares ${planPage.motionIntent.units.length} motion units but PPJ contains ${animations.length}.`, "error", { pageId: planPage.id, slide: index + 1 }));
      const actualTransition = page.transition?.type ?? "none";
      if ((planPage.motionIntent.transition ?? "none") !== actualTransition) issues.push(issue("motionTransitionMismatch", `Page ${planPage.id} transition does not match its motion intent.`, "error", { pageId: planPage.id, slide: index + 1 }));
    }
    if (page.morph) morphPairs.push({ slide: index + 1, pageId: page.id, pairs: page.morph.pairs?.length || 0 });
    if (animations.length > 4) issues.push(issue("excessiveMotionUnits", `Page ${page.id || index + 1} contains ${animations.length} motion units.`, "warning", { slide: index + 1 }));
  }
  return section(issues, { planSha256: normalized?.sha256 || null, playbackEvidence, animationCount, motionUnits, morphPairs });
}

export async function reviewPpjArtifact(bytes, options = {}) {
  const outputSha256 = sha256(bytes);
  let receipt = options.ppjReceipt;
  if (receipt && receipt.outputSha256 && receipt.outputSha256 !== outputSha256) {
    throw new Error("PPJ review receipt does not describe the supplied PPTX bytes.");
  }
  receipt ??= await projectPptxToPpj(bytes, {
    sourceUri: "review/source.pptx",
    assetRootUri: "review/assets",
    includeNodeMap: true,
    limits: options.limits,
  });
  const program = parseProgram(receipt);
  const records = programRecords(program);
  const diagnostics = nativeIssues(receipt);
  const inspection = semanticInspection(program, records, receipt, options.maxChars ?? 20_000);
  const semantic = section([...diagnostics], {
    recordCounts: recordCounts(program, records, receipt),
    inspection: { ndjson: inspection.text, truncated: inspection.truncated },
    verification: { ndjson: diagnostics.map((entry) => JSON.stringify(entry)).join("\n"), truncated: false },
  });
  const structural = section([...diagnostics], {
    summary: {
      programSha256: receipt.programSha256,
      sourceSha256: receipt.sourceSha256 || outputSha256,
      outputSha256,
      sourceBound: receipt.sourceBound,
      restoredEmbeddedProgram: receipt.restoredEmbeddedProgram,
      expandedElementCount: receipt.expandedElementCount,
      nodeMapSha256: receipt.nodeMapJson?.byteLength ? sha256(receipt.nodeMapJson) : null,
    },
    ndjson: "",
    truncated: false,
  });
  const layout = options.layout === false ? { status: "skipped", ok: true, issues: [], scope: "none" } : layoutReview(program, records, options.layoutOptions);
  const accessibility = accessibilityReview(program, records);
  const design = designReview(program, records, options);
  const motion = motionReview(program, options);
  return { semantic, structural, layout, accessibility, design, motion, receipt, program };
}
