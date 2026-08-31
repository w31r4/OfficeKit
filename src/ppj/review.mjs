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

function layoutReview(program, records) {
  const issues = [];
  const width = Number(program.design?.canvas?.width ?? program.design?.width);
  const height = Number(program.design?.canvas?.height ?? program.design?.height);
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
    if (Number.isFinite(width) && Number.isFinite(height) && (x < 0 || y < 0 || x + w > width || y + h > height)) {
      issues.push(issue("frameOutsideCanvas", `Element ${element.id || "<anonymous>"} on page ${page} extends beyond the PPJ canvas.`, "warning", { slide: page, id: element.id }));
    }
  }
  return section(issues, { scope: "PPJ frames and NativeAOT projection" });
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
  if (options.authoringPlan == null && options.changedPageIds == null) {
    return { status: "not-requested", ok: true, planSha256: null, changedPageIds: [], issues: [] };
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
  const layout = options.layout === false ? { status: "skipped", ok: true, issues: [], scope: "none" } : layoutReview(program, records);
  const design = designReview(program, records, options);
  const motion = motionReview(program, options);
  return { semantic, structural, layout, design, motion, receipt, program };
}
