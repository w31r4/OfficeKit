#!/usr/bin/env node

import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

import { FileBlob, PresentationFile } from "../src/index.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const EVIDENCE_SCHEMA = "office-kit/pptx-template-conditioned-generation-evidence/v1";
const MAX_SLIDES = 64;
const TOPOLOGY_PARTS = new Set(["[Content_Types].xml", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"]);

export const TEMPLATE_GENERATION_SOURCES = Object.freeze([
  {
    id: "suanzhi-future-2026",
    fileName: "b34ddad8cf8b_012_算秩未来2026_0127_极致技术&长期主义.pptx",
    sourceSha256: "b34ddad8cf8bbd083b60e07f8488267b1a0e4199db422468faa0eeb5d83e1762",
    minimumGeneratedSlides: 10,
    content: ["架构验证", "交付证据", "平台能力", "客户价值", "阶段成果", "风险收敛", "下一步计划", "质量门禁", "复盘结论", "行动清单"],
  },
  {
    id: "blue-gray-acid-template",
    fileName: "template.pptx",
    sourceSha256: "558ce85c0d64cd2a06faf88d6a4aa331e8cd4c685c59101c835ded2fbc87696d",
    minimumGeneratedSlides: 10,
    content: ["产品洞察", "关键动作", "用户证据", "方案结构", "数据结果", "问题拆解", "路线安排", "复盘要点", "决策事项", "工作清单"],
  },
  {
    id: "mckinsey-customer-loyalty",
    fileName: "ppt169_麦肯锡风_kimsoong_customer_loyalty.pptx",
    sourceSha256: "e0bfb89454f51c400ac03797c255aa93919328ff8dba36fe414e5bcfed0536c5",
    minimumGeneratedSlides: 10,
    content: ["Retention insight", "Customer signal", "Root cause", "Priority action", "Evidence map", "Strategy choice", "Pilot plan", "Success metric", "Decision log", "Next action"],
  },
]);

function sha256(value) {
  return crypto.createHash("sha256").update(value).digest("hex");
}

function parseArgs(argv) {
  const result = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (token === "--help" || token === "-h") result.help = true;
    else if (token.startsWith("--")) {
      const key = token.slice(2);
      const next = argv[index + 1];
      result[key] = next && !next.startsWith("--") ? argv[++index] : true;
    }
  }
  return result;
}

function usage() {
  return [
    "Usage:",
    "  officekit run scripts/pptx-template-generation.mjs --assets-dir <dir> --output-dir <dir>",
    "",
    "Imports the three frozen external PPTX samples, derives a bounded design profile,",
    "duplicates source slides, reimports the starter, applies run/SVG-level content",
    "changes, reimports again, and writes generation evidence plus a montage preview.",
    "No raw OOXML or alternate Office writer is used.",
    "",
    "Options:",
    "  --source <id>       Run one frozen source only.",
    "  --slides <n>        Generated slide count (default: 10; max: 64).",
  ].join("\n");
}

function sourceDefinition(id) {
  const definition = TEMPLATE_GENERATION_SOURCES.find((source) => source.id === id);
  if (!definition) throw new Error(`Unknown template-generation source: ${id}`);
  return definition;
}

function nonEmptyRuns(shape) {
  return (shape?.text?.paragraphs || []).flatMap((paragraph, paragraphIndex) =>
    (paragraph.runs || []).map((run, runIndex) => ({
      text: String(run.text ?? ""),
      paragraphIndex,
      runIndex,
    })).filter((run) => run.text.trim()),
  );
}

function isPlaceholderText(value) {
  const text = String(value || "").trim();
  return !text || /^click to add|^单击此处|^添加(?:副标题|小标题|正文)|^lorem ipsum/i.test(text);
}

function targetForSlide(slide) {
  const shapeCandidates = [];
  for (let shapeIndex = 0; shapeIndex < (slide.shapes?.items || []).length; shapeIndex += 1) {
    const shape = slide.shapes.items[shapeIndex];
    for (const run of nonEmptyRuns(shape)) {
      if (isPlaceholderText(run.text) && shapeCandidates.length > 0) continue;
      shapeCandidates.push({
        kind: "shape-run",
        shapeIndex,
        name: shape.name || undefined,
        runText: run.text,
        paragraphIndex: run.paragraphIndex,
        runIndex: run.runIndex,
      });
    }
  }
  const preferredShape = shapeCandidates
    .filter((candidate) => candidate.runText.trim().length >= 3)
    .sort((left, right) => right.runText.trim().length - left.runText.trim().length || left.shapeIndex - right.shapeIndex)[0];
  if (preferredShape) return preferredShape;
  for (let imageIndex = 0; imageIndex < (slide.images?.items || []).length; imageIndex += 1) {
    const image = slide.images.items[imageIndex];
    if (image.svgTextCapability?.supported !== true) continue;
    const node = image.getSvgTextNodes().find((candidate) => !isPlaceholderText(candidate.text) && candidate.text.trim().length >= 3)
      || image.getSvgTextNodes()[0];
    if (node) return { kind: "svg-text", imageIndex, nodeId: node.id, expectedHash: node.expectedHash, runText: node.text, name: image.name || undefined };
  }
  throw new Error(`No bounded text target was found on source slide ${slide.index + 1}.`);
}

function selectSourceSlides(presentation, profile, count) {
  const archetypes = Array.isArray(profile.slideArchetypes) ? profile.slideArchetypes : [];
  const candidates = archetypes
    .map((archetype) => ({ slide: Number(archetype.slide), signature: archetype.signature, textChars: Number(archetype.textChars || 0) }))
    .filter((candidate) => Number.isInteger(candidate.slide) && candidate.slide > 0)
    .filter((candidate) => presentation.slides.items[candidate.slide - 1]?.cloneCapability?.supported === true)
    .sort((left, right) => right.textChars - left.textChars || left.slide - right.slide);
  const selected = [];
  const signatures = new Set();
  for (const candidate of candidates) {
    if (selected.length >= count) break;
    if (signatures.has(candidate.signature)) continue;
    try {
      const target = targetForSlide(presentation.slides.items[candidate.slide - 1]);
      selected.push({ sourceSlide: candidate.slide, sourceSlideId: presentation.slides.items[candidate.slide - 1].id, target });
      signatures.add(candidate.signature);
    } catch {
      // A slide without a bounded text/SVG leaf is not a generation candidate.
    }
  }
  const fallback = candidates.length
    ? candidates
    : presentation.slides.items
      .map((slide, index) => ({ slide, index }))
      .filter(({ slide }) => slide.cloneCapability?.supported === true)
      .map(({ index }) => ({ slide: index + 1, signature: `slide:${index + 1}` }));
  let cursor = 0;
  while (selected.length < count) {
    const candidate = fallback[cursor % fallback.length];
    cursor += 1;
    const sourceSlide = Number(candidate.slide);
    const existing = selected.find((entry) => entry.sourceSlide === sourceSlide && entry.target);
    if (!existing) {
      try {
        const source = presentation.slides.items[sourceSlide - 1];
        selected.push({ sourceSlide, sourceSlideId: source.id, target: targetForSlide(source) });
      } catch {
        if (cursor > fallback.length * 3) break;
      }
    } else {
      selected.push({ sourceSlide, sourceSlideId: existing.sourceSlideId, target: existing.target });
    }
  }
  if (selected.length < count) throw new Error(`Only ${selected.length} source slides expose bounded text/SVG targets; ${count} are required.`);
  return selected.slice(0, count);
}

function findShapeTarget(slide, target) {
  const candidates = (slide.shapes?.items || []).filter((shape) => {
    if (target.name && shape.name !== target.name) return false;
    return nonEmptyRuns(shape).some((run) => run.text === target.runText);
  });
  const shape = candidates[0] || slide.shapes?.items?.[target.shapeIndex];
  if (!shape) throw new Error(`Cloned slide ${slide.index + 1} no longer exposes shape target.`);
  const run = nonEmptyRuns(shape).find((candidate) => candidate.text === target.runText);
  if (!run) throw new Error(`Cloned slide ${slide.index + 1} no longer exposes the expected text run.`);
  return { shape, run };
}

function applyTarget(slide, target, value) {
  if (target.kind === "svg-text") {
    const image = slide.images?.items?.[target.imageIndex];
    if (!image || image.svgTextCapability?.supported !== true) throw new Error(`Cloned slide ${slide.index + 1} no longer exposes the SVG target.`);
    image.editSvgText(target.nodeId, { expectedHash: target.expectedHash, value });
    return { kind: target.kind, value };
  }
  const { shape, run } = findShapeTarget(slide, target);
  // TextFrame.replace is deliberately run-scoped. It keeps the imported
  // paragraph/run topology instead of rebuilding a source-bound text body.
  shape.text.replace(run.text, value);
  return { kind: target.kind, value };
}

async function duplicateSourceSlidesInRounds(presentation, plan) {
  let working = presentation;
  let pending = plan.map((entry) => ({ ...entry }));
  const clones = [];
  let round = 0;
  while (pending.length) {
    round += 1;
    const usedOrigins = new Set();
    const nextPending = [];
    // Descending source order keeps original slide parts at their logical
    // package paths when the codec inserts a clone immediately after its
    // origin. The generated plan order is restored later from the locators.
    for (const entry of [...pending].sort((left, right) => right.sourceSlide - left.sourceSlide)) {
      if (usedOrigins.has(entry.sourceSlide)) {
        nextPending.push(entry);
        continue;
      }
      // Public slide ids are display-position ids and are regenerated on
      // every import. Locate the original source by its frozen ordinal plus
      // the already inserted clones from lower source ordinals instead.
      const lowerSourceClones = clones.filter((candidate) => candidate.sourceSlide < entry.sourceSlide).length;
      const sourceSlide = working.slides.items[entry.sourceSlide - 1 + lowerSourceClones];
      if (!sourceSlide?.cloneCapability?.supported) {
        throw new Error(`Source slide ${entry.sourceSlide} cannot be safely cloned: ${sourceSlide?.cloneCapability?.blockedReason || "unknown reason"}`);
      }
      const clone = sourceSlide.duplicate();
      const occurrence = clones.filter((candidate) => candidate.sourceSlide === entry.sourceSlide).length + 1;
      clones.push({ ...entry, round, occurrence });
      usedOrigins.add(entry.sourceSlide);
    }
    pending = nextPending;
    if (pending.length) {
      const staged = await PresentationFile.exportPptx(working);
      working = await PresentationFile.importPptx(staged.bytes);
    }
  }
  return { presentation: working, clones, roundCount: round };
}

function findCloneSlide(presentation, entry, allClones) {
  const lowerSourceClones = allClones.filter((candidate) => candidate.sourceSlide < entry.sourceSlide).length;
  const sourceIndex = entry.sourceSlide - 1 + lowerSourceClones;
  return presentation.slides.items[sourceIndex + entry.occurrence];
}

function targetValue(slide, target) {
  if (target.kind === "svg-text") {
    const image = slide.images?.items?.[target.imageIndex];
    return image?.getSvgTextNodes?.().find((node) => node.id === target.nodeId)?.text;
  }
  const { shape } = findShapeTarget(slide, { ...target, runText: target.value });
  return shape.text.value;
}

async function packageMap(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const result = new Map();
  for (const name of Object.keys(zip.files).sort()) {
    const entry = zip.file(name);
    if (!entry || entry.dir) continue;
    result.set(name, Buffer.from(await entry.async("nodebuffer")));
  }
  return result;
}

function xmlAttribute(attributes, name) {
  const match = String(attributes || "").match(new RegExp(`(?:^|\\s)${name}="([^"]*)"`, "u"));
  return match?.[1];
}

function issueSignature(issue) {
  return JSON.stringify({
    kind: issue?.kind,
    type: issue?.type,
    severity: issue?.severity,
    name: issue?.name,
    names: issue?.names,
  });
}

function novelIssues(sourceIssues, outputIssues) {
  const baseline = new Set((sourceIssues || []).map(issueSignature));
  const seen = new Set();
  return (outputIssues || []).filter((issue) => {
    const signature = issueSignature(issue);
    if (baseline.has(signature) || seen.has(signature)) return false;
    seen.add(signature);
    return true;
  });
}

function resolvePackageTarget(sourcePart, target) {
  const base = path.posix.dirname(sourcePart);
  const resolved = path.posix.normalize(String(target).startsWith("/") ? String(target).slice(1) : path.posix.join(base, target));
  return resolved.startsWith("ppt/") ? resolved : `ppt/${resolved.replace(/^\/?/, "")}`;
}

async function presentationSlideMap(packageParts) {
  const presentation = packageParts.get("ppt/presentation.xml")?.toString("utf8") || "";
  const relationships = packageParts.get("ppt/_rels/presentation.xml.rels")?.toString("utf8") || "";
  const targets = new Map([...relationships.matchAll(/<Relationship\b([^>]*)\/?\s*>/gu)].map((match) => [
    xmlAttribute(match[1], "Id"),
    xmlAttribute(match[1], "Target"),
  ]));
  const result = new Map();
  for (const match of presentation.matchAll(/<p:sldId\b([^>]*)\/?\s*>/gu)) {
    const id = xmlAttribute(match[1], "id");
    const relationId = xmlAttribute(match[1], "r:id");
    const target = targets.get(relationId);
    if (!id || !target) continue;
    const part = resolvePackageTarget("ppt/presentation.xml", target);
    const relPath = part.replace(/^ppt\/(.+)\/([^/]+)\.xml$/u, "ppt/$1/_rels/$2.xml.rels");
    result.set(id, { part, relPath });
  }
  return result;
}

async function packageDiff(sourceBytes, outputBytes) {
  const source = await packageMap(sourceBytes);
  const output = await packageMap(outputBytes);
  const sourceSlides = await presentationSlideMap(source);
  const outputSlides = await presentationSlideMap(output);
  const slideParts = new Set([...sourceSlides.values(), ...outputSlides.values()].flatMap((slide) => [slide.part, slide.relPath]));
  const logicalSlideMismatches = [];
  const changedExistingParts = [];
  const missingParts = [];
  for (const [name, sourcePart] of source) {
    if (slideParts.has(name)) continue;
    const outputPart = output.get(name);
    if (!outputPart) {
      missingParts.push(name);
      continue;
    }
    if (!sourcePart.equals(outputPart)) changedExistingParts.push(name);
  }
  for (const [slideId, sourceSlide] of sourceSlides) {
    const outputSlide = outputSlides.get(slideId);
    if (!outputSlide) {
      logicalSlideMismatches.push(`${slideId}:missing`);
      continue;
    }
    if (!source.get(sourceSlide.part)?.equals(output.get(outputSlide.part))) logicalSlideMismatches.push(`${slideId}:slide`);
    const sourceRels = source.get(sourceSlide.relPath);
    const outputRels = output.get(outputSlide.relPath);
    if (Boolean(sourceRels) !== Boolean(outputRels) || (sourceRels && !sourceRels.equals(outputRels))) logicalSlideMismatches.push(`${slideId}:relationships`);
  }
  const addedParts = [...output.keys()].filter((name) => !source.has(name));
  return {
    changedExistingParts,
    logicalSlideMismatches,
    nonTopologyChangedParts: [...changedExistingParts.filter((name) => !TOPOLOGY_PARTS.has(name)), ...logicalSlideMismatches],
    topologyChangedParts: changedExistingParts.filter((name) => TOPOLOGY_PARTS.has(name)),
    addedParts,
    missingParts: missingParts.filter((name) => !slideParts.has(name)),
  };
}

async function runOne(definition, assetsDir, outputDir, requestedSlides) {
  const sourcePath = path.join(assetsDir, definition.fileName);
  const sourceBytes = await fs.readFile(sourcePath);
  const sourceSha256 = sha256(sourceBytes);
  if (sourceSha256 !== definition.sourceSha256) throw new Error(`${definition.id} source SHA-256 mismatch: ${sourceSha256}`);
  const source = await PresentationFile.importPptx(new FileBlob(sourceBytes, { type: PPTX_MIME, name: definition.fileName }));
  const sourceSlideCount = source.slides.count;
  const profile = source.designProfile({ maxItems: 64 });
  const sourceLayout = source.validateLayout({ maxChars: Infinity });
  const sourceVerify = source.verify({ maxChars: Infinity });
  const generatedSlides = Math.max(definition.minimumGeneratedSlides, requestedSlides || definition.minimumGeneratedSlides);
  if (!Number.isInteger(generatedSlides) || generatedSlides < 1 || generatedSlides > MAX_SLIDES) throw new RangeError("generated slide count must be an integer from 1 through 64");
  const plan = selectSourceSlides(source, profile, generatedSlides);
  const cloned = await duplicateSourceSlidesInRounds(source, plan);
  const clones = cloned.clones;
  const starter = await PresentationFile.exportPptx(cloned.presentation);
  const starterPresentation = await PresentationFile.importPptx(starter.bytes);
  const applied = [];
  for (let index = 0; index < clones.length; index += 1) {
    const entry = clones[index];
    const slide = findCloneSlide(starterPresentation, entry, clones);
    if (!slide) throw new Error(`Generated clone ${index + 1} disappeared during starter reimport.`);
    const value = definition.content[index % definition.content.length];
    applied.push({ outputSlide: slide.index + 1, sourceSlide: entry.sourceSlide, target: entry.target, value, result: applyTarget(slide, entry.target, value) });
  }
  const output = await PresentationFile.exportPptx(starterPresentation);
  const reopened = await PresentationFile.importPptx(output.bytes);
  const verified = applied.map((entry) => {
    const slide = reopened.slides.items[entry.outputSlide - 1];
    const value = entry.target.kind === "svg-text"
      ? slide?.images?.items?.[entry.target.imageIndex]?.getSvgTextNodes?.().find((node) => node.id === entry.target.nodeId)?.text
      : findShapeTarget(slide, { ...entry.target, runText: entry.value }).shape.text.value;
    return { outputSlide: entry.outputSlide, sourceSlide: entry.sourceSlide, targetKind: entry.target.kind, expected: entry.value, actual: value, passed: typeof value === "string" && value.includes(entry.value) };
  });
  let montage;
  let visualReview = "unavailable";
  let montageError;
  try {
    montage = await reopened.export({ format: "montage", columns: 4, thumbWidth: 260 });
    visualReview = "montage-generated";
  } catch (error) {
    // An imported custom geometry can be structurally safe to clone while the
    // bounded renderer cannot yet lower its path list. Keep the generated
    // PPTX and report the visual gap explicitly; never replace it with a
    // fake preview or silently flatten the geometry.
    montageError = error instanceof Error ? error.message : String(error);
  }
  const diff = await packageDiff(sourceBytes, output.bytes);
  const layout = reopened.validateLayout({ maxChars: Infinity });
  const verify = reopened.verify({ maxChars: Infinity });
  const baselineIssues = [...sourceLayout.issues, ...sourceVerify.issues];
  const outputIssues = [...layout.issues, ...verify.issues];
  const newIssues = novelIssues(baselineIssues, outputIssues);
  const evidence = {
    id: definition.id,
    sourceFileName: definition.fileName,
    sourceSha256,
    sourceSlides: sourceSlideCount,
    generatedSlides,
    outputSlides: reopened.slides.count,
    profile,
    selection: applied.map(({ outputSlide, sourceSlide, target, value }) => ({ outputSlide, sourceSlide, target, value })),
    verification: {
      allTargetsRoundTrip: verified.every((entry) => entry.passed),
      targets: verified,
      verifyOk: verify.ok,
      verifyIssues: verify.issues,
      layoutOk: layout.ok,
      layoutIssues: layout.issues,
      sourceBaseline: {
        verifyOk: sourceVerify.ok,
        verifyIssues: sourceVerify.issues,
        layoutOk: sourceLayout.ok,
        layoutIssues: sourceLayout.issues,
      },
      newIssues,
      noNewIssues: newIssues.length === 0,
      visualReview,
      ...(montageError ? { montageError } : {}),
    },
    packageOracle: {
      sourceProtected: sha256(sourceBytes) === sourceSha256,
      ...diff,
      nonTargetPartsPreserved: diff.nonTopologyChangedParts.length === 0 && diff.missingParts.length === 0,
    },
    starterSha256: sha256(starter.bytes),
    outputSha256: sha256(output.bytes),
    outputBytes: output.bytes.length,
    montageBytes: montage?.bytes?.length ?? 0,
  };
  await fs.mkdir(outputDir, { recursive: true });
  await fs.writeFile(path.join(outputDir, `${definition.id}.pptx`), output.bytes, { flag: "w" });
  if (montage?.bytes) await fs.writeFile(path.join(outputDir, `${definition.id}.montage.svg`), montage.bytes, { flag: "w" });
  return evidence;
}

export async function runTemplateConditionedGeneration({ assetsDir, outputDir, sourceId, generatedSlides, definitions: providedDefinitions }) {
  const definitions = providedDefinitions
    ? [...providedDefinitions]
    : sourceId ? [sourceDefinition(sourceId)] : TEMPLATE_GENERATION_SOURCES;
  const sources = [];
  for (const definition of definitions) sources.push(await runOne(definition, assetsDir, outputDir, generatedSlides));
  return { schema: EVIDENCE_SCHEMA, generatedAt: new Date(0).toISOString(), sources };
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const args = parseArgs(process.argv.slice(2));
  if (args.help || !args["assets-dir"] || !args["output-dir"]) {
    console.log(usage());
    if (!args.help) process.exitCode = 2;
  } else {
    const evidence = await runTemplateConditionedGeneration({
      assetsDir: path.resolve(args["assets-dir"]),
      outputDir: path.resolve(args["output-dir"]),
      sourceId: typeof args.source === "string" ? args.source : undefined,
      generatedSlides: args.slides === undefined ? undefined : Number(args.slides),
    });
    await fs.writeFile(path.join(path.resolve(args["output-dir"]), "evidence.json"), `${JSON.stringify(evidence, null, 2)}\n`, "utf8");
    console.log(JSON.stringify({ schema: evidence.schema, sources: evidence.sources.map((source) => ({ id: source.id, generatedSlides: source.generatedSlides, allTargetsRoundTrip: source.verification.allTargetsRoundTrip, nonTargetPartsPreserved: source.packageOracle.nonTargetPartsPreserved })) }, null, 2));
  }
}
