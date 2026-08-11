import { ndjson } from "../shared/inspection.mjs";

function locator(record) {
  return { sheet: record.sheet, id: record.id, objectKind: record.kind, ...(record.name ? { name: record.name } : {}) };
}

function issue(record, type, message) {
  return { kind: "accessibilityIssue", artifactKind: "workbook", type, severity: "error", message, ...locator(record) };
}

function manualCheck(type, message, details = {}) {
  return { kind: "accessibilityManualCheck", artifactKind: "workbook", type, severity: "manual", message, ...details };
}

export function auditSpreadsheetAccessibility(records, options = {}) {
  if (!Array.isArray(records)) throw new TypeError("Workbook accessibility audit records must be an array.");
  if (!options || typeof options !== "object" || Array.isArray(options)) throw new TypeError("Workbook accessibility audit options must be an object.");
  const issues = [];
  const manualChecks = [];
  const summary = { sheets: 0, drawings: 0, meaningfulDrawings: 0, decorativeDrawings: 0, unclassifiedDrawings: 0, missingTextDrawings: 0 };
  const sheetNames = new Set();
  const meaningfulBySheet = new Map();

  for (const record of records) {
    if (!record || typeof record !== "object" || Array.isArray(record)) throw new TypeError("Workbook accessibility audit record must be an object.");
    sheetNames.add(record.sheet);
    summary.drawings += 1;
    const accessibility = record.accessibility;
    if (accessibility?.decorative === true) {
      summary.decorativeDrawings += 1;
      continue;
    }
    const objectIds = meaningfulBySheet.get(record.sheet) || [];
    meaningfulBySheet.set(record.sheet, objectIds);
    objectIds.push(record.id);
    if (!accessibility) {
      summary.unclassifiedDrawings += 1;
      issues.push(issue(record, "unclassifiedDrawing", `${record.kind} ${record.name || record.id} on worksheet ${record.sheet} has neither accessibility metadata nor an explicit decorative classification.`));
      continue;
    }
    summary.meaningfulDrawings += 1;
    if (accessibility.title === undefined && accessibility.description === undefined) {
      summary.missingTextDrawings += 1;
      issues.push(issue(record, "meaningfulDrawingTextMissing", `${record.kind} ${record.name || record.id} on worksheet ${record.sheet} is explicitly meaningful but has neither an accessibility title nor description.`));
    }
  }
  const sheetCount = options.sheetCount ?? sheetNames.size;
  if (!Number.isSafeInteger(sheetCount) || sheetCount < 0) throw new RangeError("Workbook accessibility audit sheetCount must be a non-negative safe integer.");
  summary.sheets = sheetCount;
  for (const [sheet, objectIds] of meaningfulBySheet) {
    if (!objectIds.length) continue;
    manualChecks.push(manualCheck("drawingReadingOrder", `Worksheet ${sheet} has ${objectIds.length} meaningful or unclassified drawing object(s) whose keyboard and assistive-technology order requires native-host review.`, { sheet, objectIds }));
  }
  if (sheetCount > 0) manualChecks.push(manualCheck("worksheetSemantics", "Worksheet names, table purpose and header intent, merged-cell navigation, color-only meaning, and native Excel Accessibility Checker results require author or host review."));
  const output = [...issues, ...manualChecks];
  return {
    kind: "workbookAccessibilityAudit",
    artifactKind: "workbook",
    machineCheckPassed: issues.length === 0,
    conformanceClaimed: false,
    manualReviewRequired: manualChecks.length > 0,
    summary,
    issues,
    manualChecks,
    ...ndjson(output, options.maxChars ?? Infinity),
  };
}
