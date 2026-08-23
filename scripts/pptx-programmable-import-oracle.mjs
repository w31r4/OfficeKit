#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import crypto from "node:crypto";
import { mkdir, readFile, readdir, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";

const MAX_PARTS = 10_000;
const MAX_PART_BYTES = 128 * 1024 * 1024;
const MAX_TOTAL_PART_BYTES = 2 * 1024 * 1024 * 1024;

export function sha256(value) {
  return crypto.createHash("sha256").update(value).digest("hex");
}

export async function readIntentDefinitions(filePath) {
  const definitions = JSON.parse(await readFile(filePath, "utf8"));
  if (definitions?.schema !== "office-kit/pptx-programmable-import-intents/v1") {
    throw new Error(`Unsupported intent definitions: ${definitions?.schema ?? "missing schema"}`);
  }
  return definitions;
}

export async function readContinuationDefinitions(filePath) {
  const definitions = JSON.parse(await readFile(filePath, "utf8"));
  if (definitions?.schema !== "office-kit/pptx-codex-continuation-tasks/v1") {
    throw new Error(`Unsupported continuation definitions: ${definitions?.schema ?? "missing schema"}`);
  }
  return definitions;
}

export function resolveIntent(definitions, sourceId, intentId) {
  const source = definitions.sources.find((candidate) => candidate.id === sourceId);
  if (!source) throw new Error(`Unknown source ${sourceId}`);
  const intent = source.intents.find((candidate) => candidate.id === intentId);
  if (!intent) throw new Error(`Unknown intent ${sourceId}/${intentId}`);
  return { source, intent };
}

export async function evaluatePackageOracle({ sourceBytes, outputBytes, source, intent }) {
  const sourceHash = sha256(sourceBytes);
  if (sourceHash !== source.sha256) throw new Error(`${source.id}: source SHA-256 mismatch: ${sourceHash}`);
  const [sourceParts, outputParts] = await Promise.all([zipParts(sourceBytes), zipParts(outputBytes)]);
  const sourceNames = [...sourceParts.keys()].sort();
  const outputNames = [...outputParts.keys()].sort();
  const addedParts = outputNames.filter((name) => !sourceParts.has(name));
  const removedParts = sourceNames.filter((name) => !outputParts.has(name));
  const changedParts = sourceNames.filter((name) => outputParts.has(name) && !sourceParts.get(name).equals(outputParts.get(name)));
  const expectedAdded = [...(intent.oracle.addedParts || [])].sort();
  const expectedChanged = [...intent.oracle.changedParts].sort();
  assertJsonEqual(removedParts, [], `${source.id}/${intent.id}: removed OPC parts`);
  assertJsonEqual(addedParts, expectedAdded, `${source.id}/${intent.id}: added OPC parts`);
  assertJsonEqual(changedParts, expectedChanged, `${source.id}/${intent.id}: changed OPC parts`);

  let targetMask;
  let relationshipOracle;
  let nestedPackage = null;
  if (intent.operation === "svg-text") {
    ({ targetMask, relationshipOracle } = evaluateSvgCopyOnWrite({ sourceParts, outputParts, intent }));
  } else {
    const masked = [];
    for (const partName of expectedChanged) {
      if (partName === intent.oracle.nestedPackage?.containerPart) continue;
      const result = maskNativeLeaf(sourceParts.get(partName), outputParts.get(partName), intent);
      if (!result.passed) throw new Error(`${source.id}/${intent.id}: target XML mask failed for ${partName}: ${result.reason}`);
      masked.push({ part: partName, method: "single-issued-token", matches: result.matches });
    }
    targetMask = { passed: true, parts: masked };
    relationshipOracle = compareRelationshipParts(sourceParts, outputParts, new Set());
    if (intent.oracle.nestedPackage) {
      nestedPackage = await evaluateNestedPackage({
        sourceBytes: sourceParts.get(intent.oracle.nestedPackage.containerPart),
        outputBytes: outputParts.get(intent.oracle.nestedPackage.containerPart),
        definition: intent.oracle.nestedPackage,
        intent,
      });
    }
  }

  const nonTargetParts = sourceNames.filter((name) => !expectedChanged.includes(name));
  const nonTargetMismatches = nonTargetParts.filter((name) => !sourceParts.get(name).equals(outputParts.get(name)));
  if (nonTargetMismatches.length) throw new Error(`${source.id}/${intent.id}: non-target OPC drift: ${nonTargetMismatches.join(", ")}`);
  return {
    sourceSha256: sourceHash,
    outputSha256: sha256(outputBytes),
    sourcePartCount: sourceNames.length,
    outputPartCount: outputNames.length,
    partSet: { removed: removedParts, added: addedParts, passed: true },
    changedParts,
    nonTargetPartsByteIdentical: true,
    relationships: relationshipOracle,
    targetMask,
    nestedPackage,
  };
}

const PRESENTATION_TOPOLOGY_PARTS = new Set([
  "[Content_Types].xml",
  "ppt/presentation.xml",
  "ppt/_rels/presentation.xml.rels",
]);

export async function evaluateContinuationPackageOracle({ sourceBytes, outputBytes, source, task }) {
  const sourceHash = sha256(sourceBytes);
  if (sourceHash !== source.sha256) throw new Error(`${source.id}: source SHA-256 mismatch: ${sourceHash}`);
  const [sourceParts, outputParts] = await Promise.all([zipParts(sourceBytes), zipParts(outputBytes)]);
  const sourceNames = [...sourceParts.keys()].sort();
  const outputNames = [...outputParts.keys()].sort();
  const removedParts = sourceNames.filter((name) => !outputParts.has(name));
  const addedParts = outputNames.filter((name) => !sourceParts.has(name));
  const changedExistingParts = sourceNames.filter((name) => outputParts.has(name) && !sourceParts.get(name).equals(outputParts.get(name)));
  assertJsonEqual(removedParts, [], `${source.id}/${task.id}: removed OPC parts`);
  const unexpectedChanged = changedExistingParts.filter((name) => !PRESENTATION_TOPOLOGY_PARTS.has(name));
  assertJsonEqual(unexpectedChanged, [], `${source.id}/${task.id}: non-target existing OPC drift`);
  for (const required of PRESENTATION_TOPOLOGY_PARTS) {
    if (!changedExistingParts.includes(required)) throw new Error(`${source.id}/${task.id}: expected topology part did not change: ${required}`);
  }

  const addedSlideParts = addedParts.filter((name) => /^ppt\/slides\/slide\d+[.]xml$/u.test(name));
  if (addedSlideParts.length !== 1) throw new Error(`${source.id}/${task.id}: expected one added slide part, observed ${addedSlideParts.length}`);
  const sourceOrder = presentationSlideOrder(sourceParts);
  const outputOrder = presentationSlideOrder(outputParts);
  if (sourceOrder.length !== source.slideCount) throw new Error(`${source.id}/${task.id}: source slide count is ${sourceOrder.length}, expected ${source.slideCount}`);
  if (outputOrder.length !== sourceOrder.length + 1) throw new Error(`${source.id}/${task.id}: output must append exactly one slide`);
  assertJsonEqual(outputOrder.slice(0, sourceOrder.length), sourceOrder, `${source.id}/${task.id}: existing presentation slide order`);
  if (task.targetPageAfterAppend !== outputOrder.length) throw new Error(`${source.id}/${task.id}: target page is not the appended page`);

  const sourcePresentationRelationships = parseRelationships(sourceParts.get("ppt/_rels/presentation.xml.rels"));
  const outputPresentationRelationships = parseRelationships(outputParts.get("ppt/_rels/presentation.xml.rels"));
  for (const relationship of sourcePresentationRelationships) {
    const current = outputPresentationRelationships.find(({ id }) => id === relationship.id);
    if (!current || JSON.stringify(current) !== JSON.stringify(relationship)) {
      throw new Error(`${source.id}/${task.id}: existing presentation relationship changed: ${relationship.id}`);
    }
  }
  const addedPresentationRelationships = outputPresentationRelationships.filter(({ id }) => !sourcePresentationRelationships.some((sourceRelationship) => sourceRelationship.id === id));
  if (addedPresentationRelationships.length !== 1 || !addedPresentationRelationships[0].type.endsWith("/slide")) {
    throw new Error(`${source.id}/${task.id}: expected one added presentation slide relationship`);
  }
  const appendedOrder = outputOrder.at(-1);
  if (appendedOrder.relationshipId !== addedPresentationRelationships[0].id) {
    throw new Error(`${source.id}/${task.id}: appended slide relationship is not the only added presentation relationship`);
  }
  const appendedSlidePart = resolveRelationshipTarget("ppt/_rels/presentation.xml.rels", addedPresentationRelationships[0].target);
  if (appendedSlidePart !== addedSlideParts[0]) throw new Error(`${source.id}/${task.id}: appended slide target mismatch`);

  const sourceSlideOrder = sourceOrder[task.sourceSlide - 1];
  if (!sourceSlideOrder) throw new Error(`${source.id}/${task.id}: source slide ${task.sourceSlide} is missing`);
  const sourceSlideRelationship = sourcePresentationRelationships.find(({ id }) => id === sourceSlideOrder.relationshipId);
  if (!sourceSlideRelationship) throw new Error(`${source.id}/${task.id}: source slide relationship is missing`);
  const sourceSlidePart = resolveRelationshipTarget("ppt/_rels/presentation.xml.rels", sourceSlideRelationship.target);
  const targetMask = evaluateClonedSlideMask({ sourceParts, outputParts, sourceSlidePart, appendedSlidePart, task });

  const sourceRelationshipParts = sourceNames.filter((name) => name.endsWith(".rels") && name !== "ppt/_rels/presentation.xml.rels");
  const changedSourceRelationships = sourceRelationshipParts.filter((name) => !sourceParts.get(name).equals(outputParts.get(name)));
  assertJsonEqual(changedSourceRelationships, [], `${source.id}/${task.id}: existing relationship drift`);
  const graph = evaluateAddedGraph({ sourceParts, outputParts, addedParts });
  evaluateContentTypesPreservation(sourceParts.get("[Content_Types].xml"), outputParts.get("[Content_Types].xml"), appendedSlidePart);

  return {
    sourceSha256: sourceHash,
    outputSha256: sha256(outputBytes),
    sourcePartCount: sourceNames.length,
    outputPartCount: outputNames.length,
    sourceSlideCount: sourceOrder.length,
    outputSlideCount: outputOrder.length,
    sourceSlidePart,
    appendedSlidePart,
    partSet: { removed: [], added: addedParts, passed: true },
    changedExistingParts,
    nonTargetExistingPartsByteIdentical: true,
    relationships: {
      passed: true,
      sourceRelationshipPartsByteIdentical: true,
      addedPresentationRelationship: addedPresentationRelationships[0],
    },
    targetMask,
    addedGraph: graph,
  };
}

function presentationSlideOrder(parts) {
  const xml = parts.get("ppt/presentation.xml")?.toString("utf8") || "";
  const records = [];
  for (const match of xml.matchAll(/<p:sldId\b([^>]*)\/?\s*>/gu)) {
    const attributes = xmlAttributes(match[1]);
    const relationshipId = attributes["r:id"];
    if (!attributes.id || !relationshipId) throw new Error("Malformed presentation slide identifier");
    records.push({ id: attributes.id, relationshipId });
  }
  if (!records.length) throw new Error("Presentation contains no slide identifiers");
  return records;
}

function evaluateClonedSlideMask({ sourceParts, outputParts, sourceSlidePart, appendedSlidePart, task }) {
  const sourceXml = sourceParts.get(sourceSlidePart)?.toString("utf8");
  const outputXml = outputParts.get(appendedSlidePart)?.toString("utf8");
  if (!sourceXml || !outputXml) throw new Error(`${task.id}: cloned slide XML is missing`);
  let maskedOutput = outputXml;
  const nativeMasks = [];
  for (const edit of task.edits.filter(({ operation }) => operation === "native-leaf")) {
    const matches = countToken(maskedOutput, xmlText(edit.value));
    if (matches !== 1) throw new Error(`${task.id}: expected one cloned-slide token for ${edit.value}, observed ${matches}`);
    maskedOutput = maskedOutput.replace(xmlText(edit.value), xmlText(edit.expected));
    nativeMasks.push({ method: "single-issued-token", before: edit.expected, after: edit.value, matches });
  }

  const sourceRelPart = relationshipPartFor(sourceSlidePart);
  const outputRelPart = relationshipPartFor(appendedSlidePart);
  const sourceRelationships = sourceParts.has(sourceRelPart) ? parseRelationships(sourceParts.get(sourceRelPart)) : [];
  const outputRelationships = outputParts.has(outputRelPart) ? parseRelationships(outputParts.get(outputRelPart)) : [];
  const sourceReferenced = referencedRelationshipIds(sourceXml);
  const outputReferenced = referencedRelationshipIds(maskedOutput);
  const relationMappings = [];
  const usedOutputIds = new Set();
  let svgMask = null;
  for (const sourceId of sourceReferenced) {
    const sourceRelationship = sourceRelationships.find(({ id }) => id === sourceId);
    if (!sourceRelationship) throw new Error(`${task.id}: source slide references missing relationship ${sourceId}`);
    const candidates = outputRelationships.filter((candidate) => outputReferenced.has(candidate.id) && !usedOutputIds.has(candidate.id) && relationshipsEquivalent({
      sourceRelationship,
      outputRelationship: candidate,
      sourceRelPart,
      outputRelPart,
      sourceParts,
      outputParts,
      task,
    }));
    if (candidates.length !== 1) throw new Error(`${task.id}: relationship ${sourceId} has ${candidates.length} clone matches`);
    const outputRelationship = candidates[0];
    usedOutputIds.add(outputRelationship.id);
    relationMappings.push({ sourceId, outputId: outputRelationship.id, type: sourceRelationship.type });
    if (sourceRelationship.type.endsWith("/image")) {
      const sourceTarget = resolveRelationshipTarget(sourceRelPart, sourceRelationship.target);
      const outputTarget = resolveRelationshipTarget(outputRelPart, outputRelationship.target);
      if (sourceTarget !== outputTarget && task.edits.some(({ operation }) => operation === "svg-text")) {
        svgMask = maskSvgContinuation(sourceParts.get(sourceTarget), outputParts.get(outputTarget), task.edits.filter(({ operation }) => operation === "svg-text"));
        svgMask = { ...svgMask, sourcePart: sourceTarget, outputPart: outputTarget };
      }
    }
  }
  maskedOutput = replaceRelationshipIds(maskedOutput, relationMappings);
  if (canonicalizeXml(maskedOutput) !== canonicalizeXml(sourceXml)) {
    throw new Error(`${task.id}: appended slide differs from its source outside declared text/SVG relationship masks`);
  }
  const expectsSvg = task.edits.some(({ operation }) => operation === "svg-text");
  if (expectsSvg && !svgMask) throw new Error(`${task.id}: edited SVG copy was not resolved from the appended slide`);
  return {
    passed: true,
    sourceSlidePart,
    appendedSlidePart,
    nativeMasks,
    svgMask,
    relationshipIdMappings: relationMappings,
  };
}

function relationshipsEquivalent({ sourceRelationship, outputRelationship, sourceRelPart, outputRelPart, sourceParts, outputParts, task }) {
  if (sourceRelationship.type !== outputRelationship.type || sourceRelationship.targetMode !== outputRelationship.targetMode) return false;
  if (sourceRelationship.targetMode === "External") return sourceRelationship.target === outputRelationship.target;
  const sourceTarget = resolveRelationshipTarget(sourceRelPart, sourceRelationship.target);
  const outputTarget = resolveRelationshipTarget(outputRelPart, outputRelationship.target);
  const sourceBytes = sourceParts.get(sourceTarget);
  const outputBytes = outputParts.get(outputTarget);
  if (!sourceBytes || !outputBytes) return false;
  if (sourceBytes.equals(outputBytes)) return true;
  if (!sourceRelationship.type.endsWith("/image")) return false;
  try {
    return maskSvgContinuation(sourceBytes, outputBytes, task.edits.filter(({ operation }) => operation === "svg-text")).passed;
  } catch {
    return false;
  }
}

function maskSvgContinuation(sourceBytes, outputBytes, edits) {
  if (!sourceBytes || !outputBytes || !edits.length) throw new Error("SVG continuation mask is incomplete");
  let masked = Buffer.from(outputBytes).toString("utf8");
  const masks = [];
  for (const edit of edits) {
    const after = xmlText(edit.value);
    const before = xmlText(edit.expected);
    const matches = countToken(masked, after);
    if (matches !== 1) throw new Error(`expected one SVG token for ${edit.value}, observed ${matches}`);
    masked = masked.replace(after, before);
    masks.push({ nodeId: edit.nodeId, before: edit.expected, after: edit.value, matches });
  }
  if (!Buffer.from(masked).equals(Buffer.from(sourceBytes))) throw new Error("SVG differs outside declared text masks");
  return { passed: true, method: "exact-svg-text", masks };
}

function replaceRelationshipIds(xml, mappings) {
  let output = xml;
  for (const [index, mapping] of mappings.entries()) {
    output = replaceAttributeToken(output, mapping.outputId, `__OFFICEKIT_REL_${index}__`);
  }
  for (const [index, mapping] of mappings.entries()) {
    output = replaceAttributeToken(output, `__OFFICEKIT_REL_${index}__`, mapping.sourceId);
  }
  return output;
}

function replaceAttributeToken(xml, before, after) {
  const pattern = new RegExp(`(["'])${escapeRegex(before)}\\1`, "gu");
  return xml.replace(pattern, (match, quote) => `${quote}${after}${quote}`);
}

function referencedRelationshipIds(xml) {
  return new Set([...String(xml).matchAll(/\br:(?:id|embed|link)\s*=\s*["']([^"']+)["']/gu)].map((match) => match[1]));
}

function relationshipPartFor(partName) {
  return path.posix.join(path.posix.dirname(partName), "_rels", `${path.posix.basename(partName)}.rels`);
}

function evaluateAddedGraph({ sourceParts, outputParts, addedParts }) {
  const inbound = new Map();
  for (const [relationshipPart, bytes] of outputParts) {
    if (!relationshipPart.endsWith(".rels")) continue;
    for (const relationship of parseRelationships(bytes)) {
      if (relationship.targetMode === "External") continue;
      const target = resolveRelationshipTarget(relationshipPart, relationship.target);
      if (!outputParts.has(target)) throw new Error(`Relationship ${relationshipPart} targets missing part ${target}`);
      inbound.set(target, (inbound.get(target) || 0) + 1);
    }
  }
  const orphaned = [];
  for (const partName of addedParts) {
    if (partName.endsWith(".rels")) {
      const owner = ownerPartForRelationshipPart(partName);
      if (!outputParts.has(owner)) orphaned.push(partName);
      continue;
    }
    if (!inbound.has(partName)) orphaned.push(partName);
  }
  if (orphaned.length) throw new Error(`Added OPC graph contains orphaned parts: ${orphaned.join(", ")}`);
  return { passed: true, addedParts, orphaned: [], inboundReferenced: true };
}

function ownerPartForRelationshipPart(partName) {
  if (partName === "_rels/.rels") return "";
  return partName.replace(/\/_rels\/([^/]+)[.]rels$/u, "/$1");
}

function evaluateContentTypesPreservation(sourceBytes, outputBytes, appendedSlidePart) {
  const source = contentTypeRecords(sourceBytes);
  const output = contentTypeRecords(outputBytes);
  for (const record of source) {
    if (!output.some((candidate) => JSON.stringify(candidate) === JSON.stringify(record))) {
      throw new Error(`Existing content-type declaration changed: ${JSON.stringify(record)}`);
    }
  }
  if (!output.some((record) => record.kind === "Override" && record.key === `/${appendedSlidePart}` && record.value.includes("presentationml.slide"))) {
    throw new Error(`Missing content-type override for ${appendedSlidePart}`);
  }
}

function contentTypeRecords(bytes) {
  const xml = Buffer.from(bytes).toString("utf8");
  const records = [];
  for (const match of xml.matchAll(/<(Default|Override)\b([^>]*)\/?\s*>/gu)) {
    const attributes = xmlAttributes(match[2]);
    records.push({
      kind: match[1],
      key: match[1] === "Default" ? attributes.Extension : attributes.PartName,
      value: attributes.ContentType,
    });
  }
  return records;
}

function xmlAttributes(value) {
  return Object.fromEntries([...String(value).matchAll(/([^\s=]+)\s*=\s*["']([^"']*)["']/gu)].map((match) => [match[1], match[2]]));
}

function escapeRegex(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
}

function evaluateSvgCopyOnWrite({ sourceParts, outputParts, intent }) {
  const sourceSvg = sourceParts.get(intent.oracle.sourceSvgPart);
  const addedPart = intent.oracle.addedParts?.[0];
  const outputSvg = outputParts.get(addedPart);
  if (!sourceSvg || !outputSvg) throw new Error(`SVG oracle is missing ${intent.oracle.sourceSvgPart} or ${addedPart}`);
  const svgMask = maskExactText(sourceSvg, outputSvg, xmlText(intent.expected), xmlText(intent.value));
  if (!svgMask.passed) throw new Error(`SVG mask failed: ${svgMask.reason}`);

  const slidePart = intent.oracle.changedParts.find((name) => /^ppt\/slides\/slide\d+[.]xml$/u.test(name));
  const relPart = intent.oracle.changedParts.find((name) => /ppt\/slides\/_rels\/slide\d+[.]xml[.]rels$/u.test(name));
  if (!slidePart || !relPart) throw new Error("SVG oracle requires one changed slide and relationship part");
  const sourceRelationships = parseRelationships(sourceParts.get(relPart));
  const outputRelationships = parseRelationships(outputParts.get(relPart));
  const sourceImage = sourceRelationships.find((relationship) => relationship.type.endsWith("/image") && resolveRelationshipTarget(relPart, relationship.target) === intent.oracle.sourceSvgPart);
  if (!sourceImage) throw new Error(`Source relationship does not target ${intent.oracle.sourceSvgPart}`);
  const addedRelationships = outputRelationships.filter((relationship) => !sourceRelationships.some((sourceRelationship) => sourceRelationship.id === relationship.id));
  if (addedRelationships.length !== 1) throw new Error(`Expected one added SVG relationship, observed ${addedRelationships.length}`);
  const [addedRelationship] = addedRelationships;
  if (!addedRelationship.type.endsWith("/image") || resolveRelationshipTarget(relPart, addedRelationship.target) !== addedPart) {
    throw new Error("Added relationship does not target the declared copy-on-write SVG");
  }
  for (const sourceRelationship of sourceRelationships) {
    const current = outputRelationships.find((relationship) => relationship.id === sourceRelationship.id);
    if (!current || JSON.stringify(current) !== JSON.stringify(sourceRelationship)) {
      throw new Error(`Existing relationship changed: ${sourceRelationship.id}`);
    }
  }
  const sourceSlide = sourceParts.get(slidePart).toString("utf8");
  const outputSlide = outputParts.get(slidePart).toString("utf8");
  if (countToken(outputSlide, addedRelationship.id) !== 1) throw new Error("Added SVG relationship is not referenced exactly once by the target slide");
  const maskedSlide = outputSlide.replace(addedRelationship.id, sourceImage.id);
  if (canonicalizeXml(maskedSlide) !== canonicalizeXml(sourceSlide)) throw new Error("Target slide differs outside the copy-on-write relationship binding");
  const otherRelationshipParts = compareRelationshipParts(sourceParts, outputParts, new Set([relPart]));
  return {
    targetMask: {
      passed: true,
      parts: [
        { part: addedPart, sourcePart: intent.oracle.sourceSvgPart, method: "exact-svg-text", matches: svgMask.matches },
        { part: slidePart, method: "canonical-relationship-id", matches: 1 },
      ],
    },
    relationshipOracle: {
      ...otherRelationshipParts,
      changed: [relPart],
      addedRelationship: addedRelationship.id,
      existingRelationshipsPreserved: true,
    },
  };
}

async function evaluateNestedPackage({ sourceBytes, outputBytes, definition, intent }) {
  const [sourceParts, outputParts] = await Promise.all([zipParts(sourceBytes), zipParts(outputBytes)]);
  const sourceNames = [...sourceParts.keys()].sort();
  const outputNames = [...outputParts.keys()].sort();
  assertJsonEqual(outputNames, sourceNames, "nested package part set");
  const changed = sourceNames.filter((name) => !sourceParts.get(name).equals(outputParts.get(name)));
  assertJsonEqual(changed, [...definition.changedParts].sort(), "nested package changed parts");
  const relationships = compareRelationshipParts(sourceParts, outputParts, new Set());
  const masks = [];
  for (const partName of changed) {
    const result = maskNativeLeaf(sourceParts.get(partName), outputParts.get(partName), intent);
    if (!result.passed) throw new Error(`nested target mask failed for ${partName}: ${result.reason}`);
    masks.push({ part: partName, method: "single-issued-token", matches: result.matches });
  }
  return { passed: true, changedParts: changed, nonTargetPartsByteIdentical: true, relationships, masks };
}

function compareRelationshipParts(sourceParts, outputParts, allowedChanged) {
  const sourceRelationshipParts = [...sourceParts.keys()].filter((name) => name.endsWith(".rels")).sort();
  const changed = [];
  for (const name of sourceRelationshipParts) {
    if (!outputParts.has(name)) throw new Error(`Missing relationship part ${name}`);
    if (!sourceParts.get(name).equals(outputParts.get(name))) changed.push(name);
  }
  const unexpected = changed.filter((name) => !allowedChanged.has(name));
  if (unexpected.length) throw new Error(`Unexpected relationship drift: ${unexpected.join(", ")}`);
  return { passed: true, sourceCount: sourceRelationshipParts.length, changed };
}

function maskNativeLeaf(sourceBytes, outputBytes, intent) {
  const before = wireValue(intent.leafKind, intent.expected);
  const after = wireValue(intent.leafKind, intent.value);
  return maskExactText(sourceBytes, outputBytes, before, after);
}

function maskExactText(sourceBytes, outputBytes, before, after) {
  const source = Buffer.from(sourceBytes);
  const output = Buffer.from(outputBytes);
  if (!before || before === after) return { passed: false, matches: 0, reason: "mask tokens are empty or equal" };
  const replacement = Buffer.from(after, "utf8");
  const original = Buffer.from(before, "utf8");
  let matches = 0;
  for (let index = output.indexOf(replacement); index >= 0; index = output.indexOf(replacement, index + 1)) {
    const masked = Buffer.concat([output.subarray(0, index), original, output.subarray(index + replacement.length)]);
    if (masked.equals(source)) matches += 1;
  }
  return matches === 1
    ? { passed: true, matches }
    : { passed: false, matches, reason: `expected exactly one reversible token footprint, observed ${matches}` };
}

function wireValue(leafKind, value) {
  if (leafKind === "fillRgb" || leafKind === "lineRgb") return String(value).replace(/^#/u, "").toUpperCase();
  return xmlText(value);
}

function xmlText(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

export function canonicalizeXml(value) {
  const input = String(value).replace(/^\uFEFF/u, "").replace(/^\s*<\?xml[^?]*\?>/u, "");
  const tokens = input.match(/<!--[\s\S]*?-->|<[^>]+>|[^<]+/gu) || [];
  const output = [];
  for (const token of tokens) {
    if (!token.startsWith("<")) {
      if (token.trim()) output.push(token);
      continue;
    }
    if (token.startsWith("<!--") || token.startsWith("</") || token.startsWith("<!") || token.startsWith("<?")) {
      output.push(token.replace(/\s+>/gu, ">"));
      continue;
    }
    const selfClosing = /\/\s*>$/u.test(token);
    const match = token.match(/^<([^\s/>]+)([\s\S]*?)(?:\/\s*>|>)$/u);
    if (!match) throw new Error(`Cannot canonicalize XML token: ${token.slice(0, 120)}`);
    const attributes = [];
    const attributePattern = /([^\s=]+)\s*=\s*("[^"]*"|'[^']*')/gu;
    for (const attribute of match[2].matchAll(attributePattern)) attributes.push([attribute[1], attribute[2]]);
    const residue = match[2].replace(attributePattern, "").trim();
    if (residue) throw new Error(`Cannot canonicalize XML attributes: ${residue.slice(0, 120)}`);
    attributes.sort(([left], [right]) => left.localeCompare(right));
    output.push(`<${match[1]}${attributes.map(([name, quoted]) => ` ${name}=${quoted}`).join("")}${selfClosing ? "/>" : ">"}`);
  }
  return output.join("");
}

function parseRelationships(bytes) {
  const xml = Buffer.from(bytes).toString("utf8").replace(/^\uFEFF/u, "");
  const records = [];
  for (const match of xml.matchAll(/<Relationship\b([^>]*)\/?\s*>/gu)) {
    const attributes = Object.fromEntries([...match[1].matchAll(/([^\s=]+)\s*=\s*["']([^"']*)["']/gu)].map((attribute) => [attribute[1], attribute[2]]));
    if (!attributes.Id || !attributes.Type || !attributes.Target) throw new Error("Malformed OPC Relationship");
    records.push({ id: attributes.Id, type: attributes.Type, target: attributes.Target, targetMode: attributes.TargetMode || null });
  }
  return records.sort((left, right) => left.id.localeCompare(right.id));
}

function resolveRelationshipTarget(relationshipPart, target) {
  if (target.startsWith("/")) return target.slice(1);
  const ownerPart = ownerPartForRelationshipPart(relationshipPart);
  const ownerDirectory = ownerPart ? path.posix.dirname(ownerPart) : "";
  return path.posix.normalize(path.posix.join(ownerDirectory, target));
}

async function zipParts(bytes) {
  const zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
  const names = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  if (names.length > MAX_PARTS) throw new Error(`OPC package exceeds ${MAX_PARTS} parts`);
  const parts = new Map();
  let total = 0;
  for (const name of names) {
    if (path.posix.isAbsolute(name) || name.split("/").includes("..")) throw new Error(`Unsafe OPC path ${name}`);
    const value = Buffer.from(await zip.files[name].async("uint8array"));
    if (value.byteLength > MAX_PART_BYTES) throw new Error(`OPC part exceeds ${MAX_PART_BYTES}: ${name}`);
    total += value.byteLength;
    if (total > MAX_TOTAL_PART_BYTES) throw new Error(`OPC package exceeds ${MAX_TOTAL_PART_BYTES} uncompressed bytes`);
    parts.set(name, value);
  }
  return parts;
}

export async function renderPresentationPages(inputPath, cacheRoot, contentSha256, tools = {}) {
  const cacheDir = path.join(cacheRoot, contentSha256);
  const manifestPath = path.join(cacheDir, "pages.json");
  try {
    const existing = JSON.parse(await readFile(manifestPath, "utf8"));
    if (existing.contentSha256 === contentSha256 && existing.pages?.length) return { ...existing, cacheHit: true };
  } catch {}
  await mkdir(cacheDir, { recursive: false });
  const workDir = path.join(cacheDir, "work");
  const profileDir = path.join(cacheDir, "profile");
  await mkdir(workDir);
  await mkdir(profileDir);
  const localInput = path.join(workDir, "presentation.pptx");
  await writeFile(localInput, await readFile(inputPath), { flag: "wx" });
  const soffice = tools.soffice || process.env.OFFICEKIT_SOFFICE || "soffice";
  const pdftoppm = tools.pdftoppm || process.env.OFFICEKIT_PDFTOPPM || "pdftoppm";
  runRequired(soffice, ["--headless", `-env:UserInstallation=${pathToFileURL(profileDir).href}`, "--convert-to", "pdf", "--outdir", workDir, localInput], "LibreOffice render");
  const pdfPath = path.join(workDir, "presentation.pdf");
  await stat(pdfPath);
  runRequired(pdftoppm, ["-png", "-r", "96", pdfPath, path.join(workDir, "page")], "Poppler raster");
  const pageFiles = (await readdir(workDir)).filter((name) => /^page-\d+[.]png$/u.test(name)).sort((left, right) => pageNumber(left) - pageNumber(right));
  if (!pageFiles.length) throw new Error("Poppler produced no PPTX pages");
  const pages = [];
  for (const file of pageFiles) pages.push({ page: pageNumber(file), sha256: sha256(await readFile(path.join(workDir, file))) });
  const manifest = {
    schema: "office-kit/pptx-programmable-import-render/v1",
    contentSha256,
    renderer: {
      soffice: versionLine(soffice, ["--version"]),
      pdftoppm: versionLine(pdftoppm, ["-v"]),
      dpi: 96,
    },
    pages,
  };
  await writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, { flag: "wx" });
  return { ...manifest, cacheHit: false };
}

export async function renderKeynotePresentationPages(inputPath, cacheRoot, contentSha256, tools = {}) {
  if (process.platform !== "darwin") throw new Error("Keynote rendering is available only on macOS");
  const cacheDir = path.join(cacheRoot, contentSha256);
  const manifestPath = path.join(cacheDir, "pages.json");
  try {
    const existing = JSON.parse(await readFile(manifestPath, "utf8"));
    if (existing.contentSha256 === contentSha256 && existing.renderer?.keynote && existing.pages?.length) {
      return { ...existing, cacheHit: true };
    }
  } catch {}
  await mkdir(cacheDir, { recursive: false });
  const workDir = path.join(cacheDir, "work");
  await mkdir(workDir);
  const localInput = path.join(workDir, "presentation.pptx");
  const pdfPath = path.join(workDir, "presentation.pdf");
  await writeFile(localInput, await readFile(inputPath), { flag: "wx" });
  const osascript = tools.osascript || "osascript";
  const pdftoppm = tools.pdftoppm || process.env.OFFICEKIT_PDFTOPPM || "pdftoppm";
  runRequired(osascript, [
    "-e", "on run argv",
    "-e", "set inFile to POSIX file (item 1 of argv)",
    "-e", "set outFile to POSIX file (item 2 of argv)",
    "-e", "tell application \"Keynote\"",
    "-e", "set theDoc to missing value",
    "-e", "try",
    "-e", "set theDoc to open inFile",
    "-e", "export theDoc to outFile as PDF",
    "-e", "close theDoc saving no",
    "-e", "on error errorMessage number errorNumber",
    "-e", "if theDoc is not missing value then close theDoc saving no",
    "-e", "error errorMessage number errorNumber",
    "-e", "end try",
    "-e", "end tell",
    "-e", "end run",
    localInput,
    pdfPath,
  ], "Keynote render");
  await stat(pdfPath);
  runRequired(pdftoppm, ["-png", "-r", "96", pdfPath, path.join(workDir, "page")], "Poppler raster");
  const pageFiles = (await readdir(workDir)).filter((name) => /^page-\d+[.]png$/u.test(name)).sort((left, right) => pageNumber(left) - pageNumber(right));
  if (!pageFiles.length) throw new Error("Poppler produced no Keynote-rendered PPTX pages");
  const pages = [];
  for (const file of pageFiles) pages.push({ page: pageNumber(file), sha256: sha256(await readFile(path.join(workDir, file))) });
  const manifest = {
    schema: "office-kit/pptx-programmable-import-render/v1",
    contentSha256,
    renderer: {
      keynote: versionLine(osascript, ["-e", "tell application \"Keynote\" to get version"]),
      pdftoppm: versionLine(pdftoppm, ["-v"]),
      dpi: 96,
    },
    pages,
  };
  await writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, { flag: "wx" });
  return { ...manifest, cacheHit: false };
}

export function inspectRenderedPages(sourceRender, outputRender, targetPage) {
  if (sourceRender.pages.length !== outputRender.pages.length) {
    throw new Error(`Rendered page count changed: ${sourceRender.pages.length} -> ${outputRender.pages.length}`);
  }
  const nonTargetMismatches = [];
  for (const sourcePage of sourceRender.pages) {
    const outputPage = outputRender.pages.find(({ page }) => page === sourcePage.page);
    if (!outputPage) throw new Error(`Rendered output is missing page ${sourcePage.page}`);
    if (sourcePage.page !== targetPage && sourcePage.sha256 !== outputPage.sha256) nonTargetMismatches.push(sourcePage.page);
  }
  if (nonTargetMismatches.length) throw new Error(`Non-target rendered pages changed: ${nonTargetMismatches.join(", ")}`);
  const before = sourceRender.pages.find(({ page }) => page === targetPage);
  const after = outputRender.pages.find(({ page }) => page === targetPage);
  if (!before || !after) throw new Error(`Target rendered page ${targetPage} is missing`);
  return {
    pageCount: sourceRender.pages.length,
    targetPage,
    targetPageChanged: before.sha256 !== after.sha256,
    nonTargetPagesPixelIdentical: true,
    nonTargetMismatches: [],
    sourcePageHashes: sourceRender.pages,
    outputPageHashes: outputRender.pages,
    outputCacheHit: outputRender.cacheHit,
  };
}

export function compareRenderedPages(sourceRender, outputRender, targetPage) {
  const inspection = inspectRenderedPages(sourceRender, outputRender, targetPage);
  if (!inspection.targetPageChanged) throw new Error(`Target rendered page ${targetPage} did not change`);
  return {
    ...inspection,
    passed: true,
  };
}

export function compareContinuationRenderedPages(sourceRender, outputRender, targetPage, sourcePage) {
  if (outputRender.pages.length !== sourceRender.pages.length + 1) {
    throw new Error(`Rendered continuation page count must increase by one: ${sourceRender.pages.length} -> ${outputRender.pages.length}`);
  }
  if (targetPage !== outputRender.pages.length) throw new Error(`Continuation target page ${targetPage} is not the appended page`);
  const nonTargetMismatches = [];
  for (const sourcePage of sourceRender.pages) {
    const outputPage = outputRender.pages.find(({ page }) => page === sourcePage.page);
    if (!outputPage) throw new Error(`Rendered output is missing original page ${sourcePage.page}`);
    if (sourcePage.sha256 !== outputPage.sha256) nonTargetMismatches.push(sourcePage.page);
  }
  if (nonTargetMismatches.length) throw new Error(`Non-target rendered pages changed: ${nonTargetMismatches.join(", ")}`);
  const target = outputRender.pages.find(({ page }) => page === targetPage);
  if (!target) throw new Error(`Appended rendered page ${targetPage} is missing`);
  const sourceTarget = sourcePage == null ? null : sourceRender.pages.find(({ page }) => page === sourcePage);
  if (sourcePage != null && !sourceTarget) throw new Error(`Cloned source page ${sourcePage} is missing`);
  if (sourceTarget && sourceTarget.sha256 === target.sha256) {
    throw new Error(`Appended rendered page ${targetPage} did not change from cloned source page ${sourcePage}`);
  }
  return {
    passed: true,
    sourcePageCount: sourceRender.pages.length,
    outputPageCount: outputRender.pages.length,
    targetPage,
    appendedTargetPresent: true,
    ...(sourceTarget ? { clonedSourcePage: sourcePage, appendedTargetChangedFromSource: true } : {}),
    nonTargetPagesPixelIdentical: true,
    nonTargetMismatches: [],
    sourcePageHashes: sourceRender.pages,
    outputPageHashes: outputRender.pages,
    outputCacheHit: outputRender.cacheHit,
  };
}

function runRequired(command, args, label) {
  const result = spawnSync(command, args, { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
  if (result.status !== 0) throw new Error(`${label} failed (${result.status}): ${(result.stderr || result.stdout || "").trim()}`);
}

function versionLine(command, args) {
  const result = spawnSync(command, args, { encoding: "utf8" });
  return String(result.stdout || result.stderr || "unknown").trim().split(/\r?\n/u)[0];
}

function pageNumber(fileName) {
  return Number(fileName.match(/(\d+)[.]png$/u)?.[1]);
}

function countToken(value, token) {
  return value.split(token).length - 1;
}

function assertJsonEqual(actual, expected, label) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`${label}: expected ${JSON.stringify(expected)}, observed ${JSON.stringify(actual)}`);
  }
}

if (import.meta.url === pathToFileURL(process.argv[1] || "").href) {
  process.stderr.write("This module is the evaluator library; use pptx-programmable-import-matrix.mjs.\n");
  process.exitCode = 2;
}
