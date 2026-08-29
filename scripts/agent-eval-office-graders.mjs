import { docxGradedCaseIds, gradeDocxCase } from "./agent-eval-docx-graders.mjs";
import { gradeSpreadsheetCase, spreadsheetGradedCaseIds } from "./agent-eval-spreadsheet-graders.mjs";

const defaultWeights = { machine: 45, visual: 25, security: 20, trace: 10 };

/**
 * Cross-format dispatch only. Each Office family owns its independent semantic
 * grader; keeping this module small prevents package-level orchestration from
 * becoming a second XLSX/DOCX parser. Presentation acceptance moved to the PPJ
 * 2.0 evidence ledger and no longer exercises the retired JavaScript authoring
 * surface through PromptBench.
 */
export async function gradeOfficeCase({ item, workspace, finalMessage, trace, weights = defaultWeights }) {
  const spreadsheet = await gradeSpreadsheetCase({ item, workspace, finalMessage, trace, weights });
  if (spreadsheet.supported) return spreadsheet;
  return gradeDocxCase({ item, workspace, finalMessage, trace, weights });
}

export const officeGradedCaseIds = new Set([
  ...spreadsheetGradedCaseIds,
  ...docxGradedCaseIds,
]);
