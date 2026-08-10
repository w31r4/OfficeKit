import { ndjson } from "../shared/inspection.mjs";

const MODELED_ACCESSIBILITY_KINDS = new Set(["shape", "connector", "groupShape", "image", "table", "chart"]);

function auditLocator(record) {
  const locator = {
    slide: record.slide,
    id: record.id,
    objectKind: record.kind === "groupShape" ? "group" : record.kind,
  };
  if (record.name !== undefined) locator.name = record.name;
  if (record.parentGroupId !== undefined) locator.parentGroupId = record.parentGroupId;
  return locator;
}

function modeledIssue(record, type, message) {
  return {
    kind: "accessibilityIssue",
    artifactKind: "presentation",
    type,
    severity: "error",
    message,
    ...auditLocator(record),
  };
}

function manualCheck(type, message, details = {}) {
  return {
    kind: "accessibilityManualCheck",
    artifactKind: "presentation",
    type,
    severity: "manual",
    message,
    ...details,
  };
}

export function auditPresentationAccessibility(records, options = {}) {
  if (!Array.isArray(records)) throw new TypeError("Presentation accessibility audit records must be an array.");
  if (!options || typeof options !== "object" || Array.isArray(options)) {
    throw new TypeError("Presentation accessibility audit options must be an object.");
  }

  const issues = [];
  const manualChecks = [];
  const readingOrderBySlide = new Map();
  const summary = {
    slides: 0,
    modeledObjects: 0,
    meaningfulObjects: 0,
    decorativeObjects: 0,
    unclassifiedObjects: 0,
    missingTextObjects: 0,
    opaqueNativeObjects: 0,
  };
  const slideNumbers = new Set();

  for (const record of records) {
    if (!record || typeof record !== "object" || Array.isArray(record)) {
      throw new TypeError("Presentation accessibility audit record must be an object.");
    }
    slideNumbers.add(record.slide);
    const readingOrderCandidates = readingOrderBySlide.get(record.slide) || [];
    readingOrderBySlide.set(record.slide, readingOrderCandidates);

    if (record.kind === "nativeObject") {
      summary.opaqueNativeObjects += 1;
      readingOrderCandidates.push(record.id);
      manualChecks.push(manualCheck(
        "opaqueObjectAccessibility",
        `Native ${record.nativeKind || "presentation"} object ${record.name || record.id} on slide ${record.slide} requires host or source-level accessibility review.`,
        {
          ...auditLocator(record),
          ...(record.nativeKind === undefined ? {} : { nativeKind: record.nativeKind }),
        },
      ));
      continue;
    }
    if (!MODELED_ACCESSIBILITY_KINDS.has(record.kind)) continue;

    summary.modeledObjects += 1;
    const accessibility = record.accessibility;
    if (accessibility?.decorative === true) {
      summary.decorativeObjects += 1;
      continue;
    }

    readingOrderCandidates.push(record.id);
    if (!accessibility) {
      summary.unclassifiedObjects += 1;
      issues.push(modeledIssue(
        record,
        "unclassifiedObject",
        `${record.kind === "groupShape" ? "Group" : record.kind} ${record.name || record.id} on slide ${record.slide} has neither accessibility metadata nor an explicit decorative classification.`,
      ));
      continue;
    }

    summary.meaningfulObjects += 1;
    if (accessibility.title === undefined && accessibility.description === undefined) {
      summary.missingTextObjects += 1;
      issues.push(modeledIssue(
        record,
        "meaningfulObjectTextMissing",
        `${record.kind === "groupShape" ? "Group" : record.kind} ${record.name || record.id} on slide ${record.slide} is explicitly meaningful but has neither accessibility title nor description.`,
      ));
    }
  }

  const slideCount = options.slideCount ?? slideNumbers.size;
  if (!Number.isSafeInteger(slideCount) || slideCount < 0) {
    throw new RangeError("Presentation accessibility audit slideCount must be a non-negative safe integer.");
  }
  summary.slides = slideCount;
  for (const [slide, objectIds] of readingOrderBySlide) {
    if (objectIds.length < 2) continue;
    manualChecks.push(manualCheck(
      "readingOrder",
      `Slide ${slide} has ${objectIds.length} non-decorative or opaque objects whose assistive reading order requires native-host review.`,
      {
        slide,
        objectIds,
        reason: "PowerPoint reading order is not an independent modeled OfficeKit metadata leaf; changing shape-tree order would also change visual z-order.",
      },
    ));
  }

  const recordsOut = [...issues, ...manualChecks];
  return {
    kind: "presentationAccessibilityAudit",
    artifactKind: "presentation",
    machineCheckPassed: issues.length === 0,
    conformanceClaimed: false,
    manualReviewRequired: manualChecks.length > 0,
    summary,
    issues,
    manualChecks,
    ...ndjson(recordsOut, options.maxChars ?? Infinity),
  };
}
