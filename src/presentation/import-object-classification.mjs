import { createHash } from "node:crypto";

const SOURCE_HASH = /^[0-9a-f]{64}$/u;
const CLASSIFICATIONS = new Set([
  "typed-editable",
  "native-leaf-editable",
  "source-derived-reusable",
  "opaque-preserved",
]);

function importObjectError(code, message) {
  const error = new Error(message);
  error.code = code;
  return error;
}

function sourceRevision(state) {
  const value = String(state?.opaqueOpc?.sourcePackage?.sha256 || state?.source?.packageSha256 || "").toLowerCase();
  if (!SOURCE_HASH.test(value) || !Array.isArray(state?.slides)) {
    throw importObjectError(
      "presentation_import_object_source_required",
      "Imported-object classification requires a trusted PPTX source revision.",
    );
  }
  return value;
}

function elementIds(wire, target = new Set()) {
  if (!wire?.id || target.has(wire.id)) return target;
  target.add(wire.id);
  if (wire.content?.case === "group") {
    for (const child of wire.content.value?.children || []) elementIds(child, target);
  }
  return target;
}

function objectKind(entry) {
  const semanticKind = String(entry?.wire?.content?.case || "");
  if (semanticKind !== "opaque") return semanticKind || "unknown";
  return String(entry?.model?.nativeKind || entry?.wire?.content?.value?.nativeKind || "opaque");
}

function typedOperations(entry) {
  const source = entry.wire.source;
  const model = entry.model;
  const nativeObject = entry.wire.content?.case === "opaque";
  const operations = [];
  if (!nativeObject && source.editable === true) operations.push("semantic-model");
  if (source.textEditable === true) operations.push("text");
  if (source.accessibilityEditable === true) operations.push("accessibility");
  if (model?.svgTextCapability?.supported === true) operations.push("svg-text");
  const svgLeafKinds = new Set((model?.svgEditCapability?.leaves || []).map((leaf) => leaf.leafKind));
  if (["svgFillRgb", "svgStrokeRgb", "svgOpacity"].some((kind) => svgLeafKinds.has(kind))) operations.push("svg-style");
  if (svgLeafKinds.has("svgTransformScalar")) operations.push("svg-transform");
  if (model?.oleWorkbook) operations.push("embedded-workbook");
  if (model?.oleOfficePackage) operations.push("embedded-office-package");
  if (model?.deletionCapability?.supported === true) operations.push("delete");
  return [...new Set(operations)].sort();
}

function issuedLeaves(entry, nativeLeafRecords) {
  const ids = elementIds(entry.wire);
  return nativeLeafRecords
    .filter((leaf) => ids.has(leaf.targetId) || (leaf.parentGroupId && ids.has(leaf.parentGroupId)))
    .sort((left, right) => left.leafId.localeCompare(right.leafId));
}

function issuedReuse(entry, componentRecords) {
  const occurrences = [];
  for (const candidate of componentRecords) {
    for (const occurrence of candidate.occurrences || []) {
      if (occurrence.targetId !== entry.wire.id || occurrence.reuseCapability?.supported !== true) continue;
      occurrences.push({ candidateId: candidate.candidateId, targetId: occurrence.targetId });
    }
  }
  return occurrences.sort((left, right) => left.candidateId.localeCompare(right.candidateId));
}

function dependencySummary(entry, leaves, reuse) {
  const model = entry.model;
  return Object.freeze({
    descendantObjects: Math.max(0, elementIds(entry.wire).size - 1),
    relationshipReferences: Array.isArray(model?.relationshipReferences) ? model.relationshipReferences.length : 0,
    preservedRelationships: Array.isArray(model?.rootRelationships) ? model.rootRelationships.length : 0,
    preservedParts: Array.isArray(model?.parts) ? model.parts.length : 0,
    nativeLeaves: leaves.length,
    reusableOccurrences: reuse.length,
    embeddedWorkbook: Boolean(model?.oleWorkbook),
    embeddedOfficePackage: Boolean(model?.oleOfficePackage),
    diagramText: Boolean(model?.diagramText),
    nativeChart: Boolean(model?._nativeChartSourceBinding?.()),
    svg: typeof model?.dataUrl === "string" && /^data:image\/svg\+xml(?:;|,)/iu.test(model.dataUrl),
  });
}

function blockedReason(entry, componentRecords) {
  const deletion = entry.model?.deletionCapability;
  if (deletion?.supported === false && deletion.blockedReason) return String(deletion.blockedReason).slice(0, 512);
  for (const candidate of componentRecords) {
    const occurrence = (candidate.occurrences || []).find((value) => value.targetId === entry.wire.id);
    if (occurrence?.reuseCapability?.supported === false && occurrence.reuseCapability.reason) {
      return String(occurrence.reuseCapability.reason).slice(0, 512);
    }
  }
  return `No current-revision typed, native-leaf, or source-derived reuse operation is issued for this ${objectKind(entry)} object.`;
}

function classificationFor(entry, operations, leaves, reuse) {
  if (operations.length) return "typed-editable";
  if (leaves.length) return "native-leaf-editable";
  if (reuse.length) return "source-derived-reusable";
  return "opaque-preserved";
}

function classificationReason(classification, operations, leaves, reuse, entry, componentRecords) {
  if (classification === "typed-editable") return `Issued typed operations: ${operations.join(", ")}.`;
  if (classification === "native-leaf-editable") {
    const kinds = [...new Set(leaves.map((leaf) => leaf.leafKind))].sort();
    return `Issued native leaf kinds: ${kinds.join(", ")}.`;
  }
  if (classification === "source-derived-reusable") {
    return `Issued source-derived reuse for ${reuse.length} component occurrence${reuse.length === 1 ? "" : "s"}.`;
  }
  return blockedReason(entry, componentRecords);
}

function assertSourceBinding(entry, slideId) {
  const source = entry?.wire?.source;
  const shapeTreeIndex = Number(source?.shapeTreeIndex);
  if (!entry?.wire?.id || !entry?.model || !Number.isSafeInteger(shapeTreeIndex) || shapeTreeIndex < 0 ||
      !SOURCE_HASH.test(String(source?.elementSha256 || "").toLowerCase()) ||
      !SOURCE_HASH.test(String(source?.semanticSha256 || "").toLowerCase())) {
    throw importObjectError(
      "presentation_import_object_binding_invalid",
      `Imported object ${entry?.wire?.id || "<unknown>"} on ${slideId} lacks a complete source binding.`,
    );
  }
  return source;
}

export function classifyImportedPresentationObjects(state, options = {}) {
  const revisionSha256 = sourceRevision(state);
  const nativeLeafRecords = Array.isArray(options.nativeLeafRecords) ? options.nativeLeafRecords : [];
  const componentRecords = Array.isArray(options.componentRecords) ? options.componentRecords : [];
  const records = [];
  const locators = new Set();
  const targetIds = new Set();
  for (const slideState of state.slides) {
    const slideId = String(slideState?.wire?.id || "");
    const slide = Number(slideState?.slide?.index) + 1;
    if (!slideId || !Number.isSafeInteger(slide) || slide < 1 || !Array.isArray(slideState?.entries)) {
      throw importObjectError("presentation_import_object_binding_invalid", "Imported slide state is incomplete.");
    }
    for (const entry of slideState.entries) {
      const source = assertSourceBinding(entry, slideId);
      const locatorKey = `${slideId}:${source.shapeTreeIndex}`;
      if (locators.has(locatorKey) || targetIds.has(entry.wire.id)) {
        throw importObjectError(
          "presentation_import_object_binding_invalid",
          `Imported object classification found a duplicate source locator or target ID on ${slideId}.`,
        );
      }
      locators.add(locatorKey);
      targetIds.add(entry.wire.id);
      const operations = typedOperations(entry);
      const leaves = issuedLeaves(entry, nativeLeafRecords);
      const reuse = issuedReuse(entry, componentRecords);
      const classification = classificationFor(entry, operations, leaves, reuse);
      if (!CLASSIFICATIONS.has(classification)) {
        throw importObjectError("presentation_import_object_binding_invalid", `Imported object ${entry.wire.id} has an invalid classification.`);
      }
      const leafKinds = [...new Set(leaves.map((leaf) => leaf.leafKind))].sort();
      const id = `io_${createHash("sha256")
        .update(`${revisionSha256}\0${slideId}\0${source.shapeTreeIndex}\0${source.elementSha256}`, "utf8")
        .digest("hex")
        .slice(0, 32)}`;
      records.push(Object.freeze({
        kind: "importObject",
        id,
        targetId: entry.wire.id,
        slide,
        objectKind: objectKind(entry),
        nativeKind: entry.wire.content?.case === "opaque" ? objectKind(entry) : undefined,
        name: entry.model.name || undefined,
        topLevel: true,
        classification,
        reason: classificationReason(classification, operations, leaves, reuse, entry, componentRecords),
        sourceRevisionSha256: revisionSha256,
        sourceLocator: Object.freeze({
          slideId,
          shapeTreeIndex: Number(source.shapeTreeIndex),
          expectedElementSha256: String(source.elementSha256).toLowerCase(),
          expectedSemanticSha256: String(source.semanticSha256).toLowerCase(),
        }),
        typedOperations: Object.freeze(operations),
        nativeLeafKinds: Object.freeze(leafKinds),
        nativeLeafCount: leaves.length,
        reuse: Object.freeze(reuse.map((value) => Object.freeze({ ...value }))),
        dependencies: dependencySummary(entry, leaves, reuse),
      }));
    }
  }
  return Object.freeze(records);
}
