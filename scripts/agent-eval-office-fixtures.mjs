import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

import {
  DocumentFile,
  DocumentModel,
  SpreadsheetFile,
  Workbook,
} from "../src/index.mjs";

export const XLSX_THREADED_REVIEW_FIXTURE = Object.freeze({
  workbookName: "reviewed-budget.xlsx",
  sheetName: "Forecast",
  address: "F19",
  root: Object.freeze({
    id: "{11111111-1111-4111-8111-111111111111}",
    personId: "{AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}",
    author: "Scenario owner",
    userId: "scenario.owner@example.com",
    date: "2026-07-17T09:00:00.000Z",
    text: "Please confirm the downside cash buffer before board circulation.",
  }),
  priorReply: Object.freeze({
    id: "{22222222-2222-4222-8222-222222222222}",
    personId: "{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}",
    author: "Risk reviewer",
    userId: "risk.reviewer@example.com",
    date: "2026-07-17T09:30:00.000Z",
    text: "Sensitivity analysis is attached to the approved planning case.",
  }),
  requestedReply: "Approved after sensitivity review",
});

// This uploaded workbook deliberately extends the supported root-plus-direct
// reply shape by one identity-bound reply-of-reply.  The product contract is
// fail-closed for this graph until a mutation can preserve the complete
// parent/person topology.
export const XLSX_NESTED_REPLY_BOUNDARY_FIXTURE = Object.freeze({
  workbookName: "reviewed-budget-nested.xlsx",
  sheetName: XLSX_THREADED_REVIEW_FIXTURE.sheetName,
  address: XLSX_THREADED_REVIEW_FIXTURE.address,
  root: XLSX_THREADED_REVIEW_FIXTURE.root,
  directReply: XLSX_THREADED_REVIEW_FIXTURE.priorReply,
  nestedReply: Object.freeze({
    id: "{33333333-3333-4333-8333-333333333333}",
    parentId: XLSX_THREADED_REVIEW_FIXTURE.priorReply.id,
    personId: "{CCCCCCCC-CCCC-4CCC-8CCC-CCCCCCCCCCCC}",
    author: "Legal reviewer",
    userId: "legal.reviewer@example.com",
    date: "2026-07-17T10:00:00.000Z",
    text: "Approved after sensitivity review.",
  }),
});

export const XLSX_GROWTH_UPDATE_FIXTURE = Object.freeze({
  workbookName: "operating-plan.xlsx",
  targetSheetName: "Forecast",
  canarySheetName: "Approved Baseline",
  growthAddress: "B9",
  marginAddress: "B10",
  originalGrowth: 0.08,
  replacementGrowth: 0.1,
  grossMargin: 0.6,
  revenueFormulas: Object.freeze([
    "=B4*(1+$B$9)",
    "=B5*(1+$B$9)",
    "=B6*(1+$B$9)",
  ]),
  revisedRevenue: Object.freeze([110, 121, 133.1]),
  canaryText: "Approved Baseline — do not modify",
});

// This fixture contains one recognized imported workbook connection plus the
// QueryTable that consumes it. The PromptBench task is deliberately narrower
// than general external-data editing: it may turn only the connection's
// explicit refresh-on-open bit off, leaving the QueryTable and every other
// connection property intact.
export const XLSX_CONNECTION_REFRESH_FIXTURE = Object.freeze({
  workbookName: "external-sales-refresh-on-open.xlsx",
  sheetName: "External Data",
  tableName: "ExternalSales",
  connectionId: 7,
  connectionName: "Fixture warehouse",
  connectionCommand: "SELECT Region, Revenue FROM Sales",
  connectionOpaqueValue: "kept",
});

// This is deliberately a native, source-bound PivotTable fixture rather than
// a hand-written OOXML package. The corresponding PromptBench slice may turn
// off only this uniquely-owned cache's explicit refresh-on-load request.
export const XLSX_PIVOT_REFRESH_FIXTURE = Object.freeze({
  workbookName: "regional-revenue-refresh-on-open.xlsx",
  sheetName: "Pivot Summary",
  pivotName: "Revenue by region",
  sourceSheetName: "Data",
  sourceRange: "Data!A1:B5",
  targetRange: "A1",
});

// This fixture is a deliberately bounded enterprise-workbook preservation
// case.  The ordinary line chart and Assumptions!B4 are the only editable
// semantic targets; the connection, QueryTable, PivotTable, sparkline,
// dynamic-array metadata, threaded comment, and disconnected native/opaque
// parts are preservation canaries rather than authoring claims.
export const XLSX_OPAQUE_ENTERPRISE_FIXTURE = Object.freeze({
  workbookName: "enterprise-plan.xlsx",
  assumptionsSheetName: "Assumptions",
  assumptionAddress: "B4",
  originalAssumption: 0.065,
  replacementAssumption: 0.07,
  dashboardSheetName: "Dashboard",
  chartName: "enterprise-line-chart",
  originalChartTitle: "Revenue Outlook",
  replacementChartTitle: "Updated Revenue Outlook",
  dataSheetName: "Data",
  tableName: "EnterpriseData",
  pivotSheetName: "Summary",
  pivotName: "Enterprise Revenue",
  dynamicArrayAddress: "G2:G4",
  connectionId: 9,
  connectionName: "Enterprise warehouse",
  connectionCommand: "SELECT Quarter, Revenue FROM EnterpriseSales",
  connectionOpaqueValue: "enterprise-keep",
  queryTableName: "Enterprise sales query",
  customPowerQueryPath: "customXml/itemEnterprisePowerQuery.xml",
  slicerPath: "xl/slicers/slicer1.xml",
  slicerCachePath: "xl/slicerCaches/slicerCache1.xml",
  comboChartPath: "xl/drawings/charts/chart2.xml",
  comboChartMarker: "Enterprise combo chart is source-owned",
  threadedCommentAddress: "B4",
  threadedCommentText: "Confirm the enterprise assumption before circulation.",
});

export const XLSX_OPERATING_PLAN_FIXTURE = Object.freeze({
  actualsPath: "spreadsheets/operating-plan/actuals.csv",
  assumptionsPath: "spreadsheets/operating-plan/assumptions.json",
  outputName: "FY27-operating-plan.xlsx",
  sheets: Object.freeze(["Sources", "Assumptions", "Forecast", "Dashboard", "Checks"]),
  scenarios: Object.freeze(["Base", "Upside", "Downside"]),
  requiredCharts: Object.freeze(["line", "pie"]),
  minimumActualMonths: 24,
  minimumForecastRows: 12,
});

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
function xmlEscape(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

export const DOCX_CLASSIC_COMMENT_FIXTURE = Object.freeze({
  documentName: "legal-review.docx",
  title: "Controlled rollout legal review",
  anchorText: "Decision: proceed with controlled rollout.",
  supportingText: "The control plan, owner, and retention schedule remain unchanged.",
  comment: Object.freeze({
    author: "Legal reviewer",
    initials: "LR",
    date: "2026-07-18T09:00:00Z",
    originalText: "Please confirm the final retention wording.",
    replacementText: "Approved after legal review.",
  }),
});

// This uploaded document deliberately extends the supported modern-comment
// root-plus-direct-reply shape with one identity-bound reply-of-reply.  The
// Documents contract is fail-closed for this graph until a mutation can
// preserve every commentsExtended/commentsIds/people relationship.
export const DOCX_MODERN_COMMENT_REPLY_BOUNDARY_FIXTURE = Object.freeze({
  documentName: "modern-comment-replies.docx",
  title: "Modern review reply boundary",
  anchorText: "Decision: preserve the bounded modern review thread.",
  root: Object.freeze({
    paraId: "11111111",
    durableId: "33333333",
    id: "0",
    author: "Lead reviewer",
    initials: "LR",
    date: "2026-07-19T08:00:00Z",
    text: "Please confirm the release evidence.",
    providerId: "provider-a",
    userId: "lead@example.test",
  }),
  directReply: Object.freeze({
    paraId: "22222222",
    durableId: "44444444",
    id: "1",
    parentParaId: "11111111",
    author: "Release reviewer",
    initials: "RR",
    date: "2026-07-19T08:05:00Z",
    text: "The evidence is attached.",
    providerId: "provider-b",
    userId: "release@example.test",
  }),
  nestedReply: Object.freeze({
    paraId: "33333333",
    durableId: "55555555",
    id: "2",
    parentParaId: "22222222",
    author: "Nested reviewer",
    initials: "NR",
    date: "2026-07-19T08:10:00Z",
    text: "Approved after legal review.",
    providerId: "provider-c",
    userId: "nested@example.test",
  }),
});

// This uploaded document deliberately combines the imported table constructs
// that the bounded table model cannot safely rewrite as one topology change:
// vertical merges, a nested table, a custom table style, a tracked cell, and
// a table-cell content control.  The PromptBench task must refuse before it
// can flatten or rebuild any of those source-owned graphs.
export const DOCX_COMPLEX_TABLE_TOPOLOGY_FIXTURE = Object.freeze({
  documentName: "clinical-form.docx",
  title: "Clinical medication review form",
  leadText: "The first table is an ordinary canary; the second table is source-bound.",
  canaryText: "Do not modify this ordinary table.",
  complexTable: Object.freeze({
    styleId: "ClinicalGrid",
    headers: Object.freeze(["Medication", "Dose", "Route", "Status"]),
    mergeRoot: "Antibiotic",
    mergeContinuation: "Antibiotic",
    revisedCell: "Pending review",
    contentControl: "Reviewed by pharmacist",
    nestedCell: "Timing: with food",
  }),
});

// This fixture is intentionally narrow: two ordinary paragraphs share one
// uniquely used default HeaderPart, while a PAGE footer is a canary for the
// source-owned field boundary. The ready PromptBench task may change only the
// first header paragraph and must leave every other package part byte-stable.
export const DOCX_HEADER_TEXT_FIXTURE = Object.freeze({
  documentName: "board-brief-header.docx",
  title: "Board brief — controlled rollout",
  body: Object.freeze([
    "Decision: proceed with the approved controls and named accountable owner.",
    "The review record, retention schedule, and approval evidence remain unchanged.",
    "This document's header is the only requested source-bound edit.",
  ]),
  header: Object.freeze({
    sectionIndex: 0,
    referenceType: "default",
    originalText: "Northwind | Internal",
    replacementText: "Northwind | Reviewed",
    companionText: "Retain the body and footer exactly.",
  }),
  footer: Object.freeze({
    text: "1",
    fieldInstruction: "PAGE",
  }),
});

// The footer companion intentionally mirrors the header fixture's narrow
// source-bound profile. A PAGE header is the field-driven canary: the ready
// PromptBench task must not turn page furniture into a broad text-rebuild API.
export const DOCX_FOOTER_TEXT_FIXTURE = Object.freeze({
  documentName: "board-brief-footer.docx",
  title: "Board brief — footer review",
  body: Object.freeze([
    "Decision: proceed with the approved controls and named accountable owner.",
    "The review record, retention schedule, and approval evidence remain unchanged.",
    "This document's footer is the only requested source-bound edit.",
  ]),
  footer: Object.freeze({
    sectionIndex: 0,
    referenceType: "default",
    originalText: "Northwind | Internal",
    replacementText: "Northwind | Reviewed",
    companionText: "Retain the header and body exactly.",
  }),
  header: Object.freeze({
    text: "1",
    fieldInstruction: "PAGE",
  }),
});

// The first modeled section owns a lower-Roman PAGE field presentation. The
// ready transaction may change only that canonical w:pgNumType leaf to decimal;
// a second modeled section and the terminal section, plus all three FooterParts,
// are independent raw-OPC and native-render canaries.
export const DOCX_SECTION_PAGE_NUMBERING_FIXTURE = Object.freeze({
  documentName: "front-matter-page-numbering.docx",
  title: "Board packet — section page numbering",
  body: Object.freeze([
    "Front-matter cover canary: only its page-number display format may change.",
    "Roman-section canary: retain this section's native numbering metadata.",
    "Body-section canary: retain decimal numbering and all page furniture.",
  ]),
  target: Object.freeze({
    blockIndex: 1,
    sectionOrdinal: 0,
    originalPageNumbering: Object.freeze({ start: 1, format: "lowerRoman" }),
    replacementPageNumbering: Object.freeze({ start: 1, format: "decimal" }),
  }),
  sibling: Object.freeze({
    blockIndex: 3,
    sectionOrdinal: 1,
    pageNumbering: Object.freeze({ start: 1, format: "upperLetter" }),
  }),
  nativeSectionCount: 3,
  footerSectionCount: 3,
  footer: Object.freeze({ text: "1", fieldInstruction: "PAGE" }),
});

// A self-authored source-bound board-review document.  The two requested
// semantic edits are deliberately ordinary paragraph/table-cell patches plus
// one classic comment text edit; tracked changes, modern comment replies,
// content controls, TOC/REF fields, footnotes, VML watermark, mixed sections,
// and first/even page furniture are preservation canaries.
export const DOCX_SURGICAL_BOARD_REVIEW_FIXTURE = Object.freeze({
  documentName: "board-review.docx",
  title: "Board review — controlled release",
  recommendation: Object.freeze({
    originalText: "Recommendation: continue the pilot.",
    replacementText: "Recommendation: approve controlled release.",
  }),
  riskTable: Object.freeze({
    headers: Object.freeze(["Risk", "Status"]),
    row: 1,
    column: 1,
    targetLabel: "Data migration",
    originalStatus: "Amber",
    replacementStatus: "Green",
  }),
  comment: Object.freeze({
    author: "Audit committee",
    initials: "AC",
    date: "2026-07-19T09:00:00Z",
    originalText: "Please confirm the final retention wording.",
    replacementText: "Confirmed by the audit committee.",
  }),
  modernComment: Object.freeze({
    rootParaId: "17777777",
    directReplyParaId: "18888888",
    nestedReplyParaId: "19999999",
  }),
  footnoteText: "Board packet evidence remains source-owned.",
  watermarkText: "CONFIDENTIAL — BOARD REVIEW",
  sectionCount: 3,
  headerTexts: Object.freeze(["Northwind | Internal", "Northwind | First page", "Northwind | Even page"]),
  complexTable: Object.freeze({
    styleId: "BoardReviewGrid",
    headers: Object.freeze(["Medication", "Dose", "Route", "Status"]),
    mergeRoot: "Antibiotic",
    mergeContinuation: "Antibiotic",
    revisedCell: "Pending review",
    contentControl: "Reviewed by pharmacist",
    nestedCell: "Timing: with food",
  }),
  refBookmark: "BoardRecommendation",
});


function commentConfig(comment) {
  return {
    id: comment.id,
    personId: comment.personId,
    author: comment.author,
    date: comment.date,
    person: {
      id: comment.personId,
      displayName: comment.author,
      userId: comment.userId,
      providerId: "None",
    },
  };
}

export async function generateXlsxThreadedReview(target) {
  const fixture = XLSX_THREADED_REVIEW_FIXTURE;
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add(fixture.sheetName);
  sheet.getRange("A1:F19").values = [
    ["FY27 downside cash review", null, null, null, null, null],
    ["Scenario", "Revenue", "Gross margin", "EBITDA", "Cash buffer", "Board status"],
    ["Base", 1480, 0.58, 255, 122, "Approved"],
    ["Upside", 1640, 0.6, 302, 168, "Approved"],
    ["Downside", 1210, 0.51, 142, 45, "Pending review"],
    [null, null, null, null, null, null],
    ["Control", "Value", null, null, null, null],
    ["Minimum required cash buffer", 40, null, null, null, null],
    ["Downside buffer check", null, null, null, null, null],
    [null, null, null, null, null, null],
    ["Board review notes", null, null, null, null, null],
    ["The threaded review target is the final board-status cell below.", null, null, null, null, null],
    [null, null, null, null, null, null],
    ["Forecast", "Value", null, null, null, null],
    ["Opening cash", 210, null, null, null, null],
    ["Operating cash flow", -118, null, null, null, null],
    ["Committed spend", -47, null, null, null, null],
    ["Minimum buffer", -40, null, null, null, null],
    ["Downside cash buffer", null, null, null, null, null],
  ];
  sheet.getRange("B9").formulas = [["=IF(F5=\"Pending review\",\"REVIEW\",\"PASS\")"]];
  sheet.getRange("F19").formulas = [["=SUM(B15:B18)"]];
  sheet.getRange("A1:F1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF", size: 14 } };
  sheet.getRange("A2:F2").format = { fill: "#E2E8F0", font: { bold: true } };
  sheet.getRange("A14:B14").format = { fill: "#E2E8F0", font: { bold: true } };
  sheet.getRange("A1:F19").format.columnWidthPx = 130;
  sheet.getRange("A1:A19").format.columnWidthPx = 220;
  sheet.freezePanes.freezeRows(2);

  workbook.comments.setSelf({ displayName: "Finance workflow" });
  const thread = workbook.comments.addThread(
    { cell: sheet.getRange(fixture.address) },
    fixture.root.text,
    { id: "downside-cash-review", author: fixture.root.author, resolved: false, comment: commentConfig(fixture.root) },
  );
  thread.addReply(fixture.priorReply.text, commentConfig(fixture.priorReply));

  workbook.recalculate();
  const exported = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await exported.arrayBuffer()));
  return { path: target, type: XLSX_MIME };
}

export async function generateXlsxNestedReplyBoundary(target) {
  const fixture = XLSX_NESTED_REPLY_BOUNDARY_FIXTURE;
  await generateXlsxThreadedReview(target);
  const zip = await JSZip.loadAsync(await fs.readFile(target));
  const personPath = "xl/persons/person.xml";
  const threadedPath = "xl/threadedcomments/threadedcomment.xml";
  const personXml = await zip.file(personPath)?.async("text");
  const threadedXml = await zip.file(threadedPath)?.async("text");
  if (!personXml || !threadedXml) throw new Error("threaded review fixture is missing canonical Office parts");
  const person = fixture.nestedReply;
  const personFragment = `<xltc:person displayName="${xmlEscape(person.author)}" id="${person.personId}" userId="${xmlEscape(person.userId)}" providerId="None" />`;
  const commentFragment = `<xltc:threadedComment ref="${fixture.address}" dT="${person.date.replace(/\.000Z$/, "Z")}" personId="${person.personId}" id="${person.id}" parentId="${person.parentId}" done="0"><xltc:text>${xmlEscape(person.text)}</xltc:text></xltc:threadedComment>`;
  if (!personXml.includes("</xltc:personList>") || !threadedXml.includes("</xltc:ThreadedComments>")) {
    throw new Error("threaded review fixture uses an unexpected Office namespace shape");
  }
  zip.file(personPath, personXml.replace("</xltc:personList>", `${personFragment}</xltc:personList>`));
  zip.file(threadedPath, threadedXml.replace("</xltc:ThreadedComments>", `${commentFragment}</xltc:ThreadedComments>`));
  // The normal model exporter intentionally allocates relationship IDs per
  // workbook.  PromptBench assets need reproducible bytes, so canonicalize
  // those package-local IDs and ZIP metadata before publishing the fixture.
  const entries = [];
  for (const name of Object.keys(zip.files).filter((entry) => !zip.files[entry].dir).sort()) {
    const file = zip.file(name);
    if (!file) continue;
    entries.push({ name, bytes: await file.async("uint8array") });
  }
  const relationshipIds = new Map();
  const xmlEntries = entries.map(({ name, bytes }) => {
    if (!/\.(?:xml|rels)$/i.test(name)) return { name, bytes, xml: null };
    const xml = new TextDecoder().decode(bytes);
    for (const match of xml.matchAll(/\bR[a-f0-9]{16}\b/gi)) {
      if (!relationshipIds.has(match[0])) relationshipIds.set(match[0], `rId${relationshipIds.size + 1}`);
    }
    return { name, bytes, xml };
  });
  const fixedDate = new Date("1980-01-01T00:00:00.000Z");
  const canonicalZip = new JSZip();
  for (const entry of xmlEntries) {
    const data = entry.xml === null
      ? entry.bytes
      : entry.xml.replace(/\bR[a-f0-9]{16}\b/gi, (id) => relationshipIds.get(id) || id);
    canonicalZip.file(entry.name, data, {
      date: fixedDate,
      createFolders: false,
      compression: "DEFLATE",
      compressionOptions: { level: 9 },
    });
  }
  const bytes = await canonicalZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 9 }, platform: "DOS" });
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, bytes);
  return { path: target, type: XLSX_MIME };
}

export async function generateXlsxGrowthUpdate(target) {
  const fixture = XLSX_GROWTH_UPDATE_FIXTURE;
  const workbook = Workbook.create();
  const forecast = workbook.worksheets.add(fixture.targetSheetName);
  forecast.getRange("A1:D10").values = [
    ["FY27 Operating Plan", null, null, null],
    ["Update only the monthly growth assumption; preserve every formula and the approved baseline.", null, null, null],
    ["Month", "Revenue", "Gross Profit", "Growth"],
    ["Jan", 100, null, null],
    ["Feb", null, null, null],
    ["Mar", null, null, null],
    ["Apr", null, null, null],
    [null, null, null, null],
    ["Monthly growth", fixture.originalGrowth, null, null],
    ["Gross margin", fixture.grossMargin, null, null],
  ];
  forecast.getRange("B5:B7").formulas = fixture.revenueFormulas.map((formula) => [formula]);
  forecast.getRange("C4:C7").formulas = [
    ["=B4*$B$10"],
    ["=B5*$B$10"],
    ["=B6*$B$10"],
    ["=B7*$B$10"],
  ];
  forecast.getRange("D4").formulas = [["=0"]];
  forecast.getRange("D5:D7").formulas = [
    ["=B5/B4-1"],
    ["=B6/B5-1"],
    ["=B7/B6-1"],
  ];
  forecast.getRange("A1:D1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF", size: 14 } };
  forecast.getRange("A3:D3").format = { fill: "#E2E8F0", font: { bold: true } };
  forecast.getRange("A9:B10").format = { fill: "#FEF3C7", font: { bold: true } };
  forecast.getRange("B4:C7").setNumberFormat("$#,##0.00");
  forecast.getRange("B9:B10").setNumberFormat("0.0%");
  forecast.getRange("D4:D7").setNumberFormat("0.0%");
  forecast.getRange("A1:D10").format.columnWidthPx = 150;
  forecast.getRange("A1:A10").format.columnWidthPx = 280;
  forecast.freezePanes.freezeRows(3);

  const baseline = workbook.worksheets.add(fixture.canarySheetName);
  baseline.getRange("A1:C5").values = [
    [fixture.canaryText, null, null],
    ["Metric", "Approved value", "Status"],
    ["Monthly growth", fixture.originalGrowth, "Board approved"],
    ["Gross margin", fixture.grossMargin, "Board approved"],
    ["Scope", "No changes authorized", "Canary"],
  ];
  baseline.getRange("A1:C1").format = { fill: "#14532D", font: { bold: true, color: "#FFFFFF", size: 14 } };
  baseline.getRange("A2:C2").format = { fill: "#DCFCE7", font: { bold: true } };
  baseline.getRange("B3:B4").setNumberFormat("0.0%");
  baseline.getRange("A1:C5").format.columnWidthPx = 170;
  baseline.getRange("A1:A5").format.columnWidthPx = 260;
  baseline.freezePanes.freezeRows(2);

  workbook.recalculate();
  const verification = workbook.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated XLSX growth-update fixture failed model verification: " + verification.ndjson);
  const exported = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await exported.arrayBuffer()));
  return { path: target, type: XLSX_MIME };
}

async function attachConnectionRefreshFixture(file, fixture) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const zip = await JSZip.loadAsync(bytes);
  const tablePartPath = Object.keys(zip.files).find((name) => /^xl\/tables\/table1\.xml$/i.test(name));
  if (!tablePartPath || zip.file("xl/connections.xml")) throw new Error("XLSX connection-refresh fixture could not find an unbound TablePart.");
  const tableRelationshipPath = `${path.posix.dirname(tablePartPath)}/_rels/${path.posix.basename(tablePartPath)}.rels`;
  const [contentTypes, workbookRelationships] = await Promise.all([
    zip.file("[Content_Types].xml")?.async("text"),
    zip.file("xl/_rels/workbook.xml.rels")?.async("text"),
  ]);
  if (!contentTypes?.includes("</Types>") || !workbookRelationships?.includes("</Relationships>") || zip.file(tableRelationshipPath)) {
    throw new Error("XLSX connection-refresh fixture could not safely add the required OPC relationships.");
  }
  const queryPartPath = "xl/queryTables/queryTable1.xml";
  zip.file("[Content_Types].xml", contentTypes.replace(
    "</Types>",
    `<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/><Override PartName="/${queryPartPath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/></Types>`,
  ));
  zip.file("xl/_rels/workbook.xml.rels", workbookRelationships.replace(
    "</Relationships>",
    '<Relationship Id="rIdConnections" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/></Relationships>',
  ));
  zip.file(tableRelationshipPath, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdQueryTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable" Target="../queryTables/queryTable1.xml"/></Relationships>');
  zip.file("xl/connections.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:connections xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:fixture="urn:office-kit:promptbench"><x:connection id="${fixture.connectionId}" name="${xmlEscape(fixture.connectionName)}" description="Read-only warehouse source" type="5" refreshedVersion="8" keepAlive="0" interval="30" background="1" refreshOnLoad="1" saveData="1" savePassword="0" credentials="integrated"><x:dbPr connection="Provider=Fixture.Provider;Data Source=fixture.invalid" command="${xmlEscape(fixture.connectionCommand)}" commandType="2"/><x:extLst><x:ext uri="{E5A74D42-D212-4CC7-9D5B-A7393F4D8A61}"><fixture:connectionOpaque value="${xmlEscape(fixture.connectionOpaqueValue)}"/></x:ext></x:extLst></x:connection></x:connections>`);
  zip.file(queryPartPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:queryTable xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:fixture="urn:office-kit:promptbench" name="Warehouse sales" headers="1" rowNumbers="0" disableRefresh="0" backgroundRefresh="1" firstBackgroundRefresh="0" refreshOnLoad="0" growShrinkType="insertClear" fillFormulas="0" removeDataOnSave="0" disableEdit="0" preserveFormatting="1" adjustColumnWidth="1" intermediate="0" connectionId="${fixture.connectionId}"><x:queryTableRefresh preserveSortFilterLayout="1" fieldIdWrapped="0" headersInLastRefresh="1" minimumVersion="0" nextId="3" unboundColumnsLeft="0" unboundColumnsRight="0"><x:queryTableFields count="2"><x:queryTableField id="1" name="Region" dataBound="1" tableColumnId="1" fillFormulas="0" clipped="0"/><x:queryTableField id="2" name="Revenue" dataBound="1" tableColumnId="2"/></x:queryTableFields></x:queryTableRefresh><x:extLst><x:ext uri="{A1D56E5F-35B8-4C51-9C80-779E6A39D52B}"><fixture:queryOpaque value="kept"/></x:ext></x:extLst></x:queryTable>`);
  return new Uint8Array(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
}

async function canonicalizeXlsxZip(zip) {
  const entries = [];
  const relationshipIds = new Map();
  for (const name of Object.keys(zip.files).filter((entry) => !zip.files[entry].dir).sort()) {
    const file = zip.file(name);
    if (!file) continue;
    const bytes = await file.async("uint8array");
    if (!/\.(?:xml|rels)$/i.test(name)) {
      entries.push({ name, bytes, xml: null });
      continue;
    }
    const xml = new TextDecoder().decode(bytes);
    for (const match of xml.matchAll(/\bR[a-f0-9]{16}\b/gi)) {
      if (!relationshipIds.has(match[0])) relationshipIds.set(match[0], `rId${relationshipIds.size + 1}`);
    }
    entries.push({ name, bytes, xml });
  }
  const fixedDate = new Date("1980-01-01T00:00:00.000Z");
  const canonicalZip = new JSZip();
  for (const entry of entries) {
    const data = entry.xml === null
      ? entry.bytes
      : entry.xml.replace(/\bR[a-f0-9]{16}\b/gi, (id) => relationshipIds.get(id) || id);
    canonicalZip.file(entry.name, data, {
      date: fixedDate,
      createFolders: false,
      compression: "DEFLATE",
      compressionOptions: { level: 9 },
    });
  }
  return new Uint8Array(await canonicalZip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
    platform: "DOS",
  }));
}

function relationshipTargetMap(xml = "") {
  const targets = new Map();
  for (const match of String(xml).matchAll(/<Relationship\b([^>]*)\/?\s*>/g)) {
    const id = /\bId="([^"]+)"/i.exec(match[1])?.[1];
    const target = /\bTarget="([^"]+)"/i.exec(match[1])?.[1];
    if (id && target) targets.set(id, target);
  }
  return targets;
}

async function enterpriseWorksheetPath(zip, sheetName) {
  const workbookXml = await zip.file("xl/workbook.xml")?.async("text") || "";
  const workbookRels = await zip.file("xl/_rels/workbook.xml.rels")?.async("text") || "";
  const targets = relationshipTargetMap(workbookRels);
  const sheet = [...workbookXml.matchAll(/<[^>]*sheet\b([^>]*)\/?\s*>/g)]
    .map((match) => ({
      name: /\bname="([^"]+)"/i.exec(match[1])?.[1] || "",
      relId: /\br:id="([^"]+)"/i.exec(match[1])?.[1] || "",
    }))
    .find((candidate) => candidate.name === sheetName);
  if (!sheet) return null;
  const target = targets.get(sheet.relId);
  if (!target) return null;
  const normalized = String(target).replace(/^\/+/, "");
  return path.posix.normalize(normalized.startsWith("xl/") ? normalized : path.posix.join("xl", normalized));
}

async function attachOpaqueEnterpriseFixture(file, fixture) {
  const zip = await JSZip.loadAsync(new Uint8Array(await file.arrayBuffer()));
  const [contentTypes, workbookRelationships, dataSheetPath] = await Promise.all([
    zip.file("[Content_Types].xml")?.async("text"),
    zip.file("xl/_rels/workbook.xml.rels")?.async("text"),
    enterpriseWorksheetPath(zip, fixture.dataSheetName),
  ]);
  const tablePartPath = Object.keys(zip.files).find((name) => /^xl\/tables\/table1\.xml$/i.test(name));
  if (!contentTypes?.includes("</Types>") || !workbookRelationships?.includes("</Relationships>") || !tablePartPath || !dataSheetPath) {
    throw new Error("XLSX opaque-enterprise fixture could not locate the canonical workbook/table/data parts.");
  }
  const queryPartPath = "xl/queryTables/queryTable1.xml";
  const tableRelationshipPath = `${path.posix.dirname(tablePartPath)}/_rels/${path.posix.basename(tablePartPath)}.rels`;
  if (zip.file(tableRelationshipPath)) throw new Error("XLSX opaque-enterprise fixture found an unexpected existing table relationship part.");
  const overrides = [
    ["/xl/connections.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"],
    [`/${queryPartPath}`, "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"],
    [`/${fixture.slicerPath}`, "application/vnd.ms-excel.slicer+xml"],
    [`/${fixture.slicerCachePath}`, "application/vnd.ms-excel.slicerCache+xml"],
    [`/${fixture.comboChartPath}`, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"],
  ];
  let updatedContentTypes = contentTypes;
  for (const [partName, contentType] of overrides) {
    if (!updatedContentTypes.includes(`PartName="${partName}"`)) {
      updatedContentTypes = updatedContentTypes.replace("</Types>", `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`);
    }
  }
  zip.file("[Content_Types].xml", updatedContentTypes);
  let updatedWorkbookRelationships = workbookRelationships;
  if (!updatedWorkbookRelationships.includes('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections"')) {
    updatedWorkbookRelationships = updatedWorkbookRelationships.replace(
      "</Relationships>",
      '<Relationship Id="rIdEnterpriseConnections" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/></Relationships>',
    );
  }
  zip.file("xl/_rels/workbook.xml.rels", updatedWorkbookRelationships);
  zip.file(tableRelationshipPath, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdEnterpriseQueryTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable" Target="../queryTables/queryTable1.xml"/></Relationships>');
  zip.file("xl/connections.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:connections xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:fixture="urn:office-kit:promptbench"><x:connection id="${fixture.connectionId}" name="${xmlEscape(fixture.connectionName)}" description="Enterprise warehouse source" type="5" refreshedVersion="8" keepAlive="0" interval="30" background="1" refreshOnLoad="1" saveData="1" savePassword="0" credentials="integrated"><x:dbPr connection="Provider=Enterprise.Fixture;Data Source=fixture.invalid" command="${xmlEscape(fixture.connectionCommand)}" commandType="2"/><x:extLst><x:ext uri="{E5A74D42-D212-4CC7-9D5B-A7393F4D8A61}"><fixture:connectionOpaque value="${xmlEscape(fixture.connectionOpaqueValue)}"/></x:ext></x:extLst></x:connection></x:connections>`);
  zip.file(queryPartPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:queryTable xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:fixture="urn:office-kit:promptbench" name="${xmlEscape(fixture.queryTableName)}" headers="1" rowNumbers="0" disableRefresh="0" backgroundRefresh="1" refreshOnLoad="0" connectionId="${fixture.connectionId}"><x:queryTableRefresh preserveSortFilterLayout="1" fieldIdWrapped="0" headersInLastRefresh="1" minimumVersion="0" nextId="3"><x:queryTableFields count="2"><x:queryTableField id="1" name="Quarter" dataBound="1" tableColumnId="1"/><x:queryTableField id="2" name="Revenue" dataBound="1" tableColumnId="2"/></x:queryTableFields></x:queryTableRefresh><x:extLst><x:ext uri="{A1D56E5F-35B8-4C51-9C80-779E6A39D52B}"><fixture:queryOpaque value="enterprise-keep"/></x:ext></x:extLst></x:queryTable>`);
  zip.file(fixture.customPowerQueryPath, `<?xml version="1.0" encoding="UTF-8"?><PowerQuery xmlns="urn:office-kit:enterprise-query"><Query name="EnterpriseSales" source="fixture.invalid" refresh="manual">${fixture.comboChartMarker}</Query></PowerQuery>`);
  zip.file(fixture.slicerPath, `<?xml version="1.0" encoding="UTF-8"?><x:slicer xmlns:x="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="EnterpriseRegion" cache="EnterpriseRegionCache"><x:extLst><x:ext uri="{4A3D1B9B-68D8-44CA-9C61-ENTERPRISE01}"><fixture:opaque xmlns:fixture="urn:office-kit:promptbench" value="slicer-keep"/></x:ext></x:extLst></x:slicer>`);
  zip.file(fixture.slicerCachePath, `<?xml version="1.0" encoding="UTF-8"?><x:slicerCache xmlns:x="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="EnterpriseRegionCache" sourceName="Region"><x:extLst><x:ext uri="{4A3D1B9B-68D8-44CA-9C61-ENTERPRISE02}"><fixture:opaque xmlns:fixture="urn:office-kit:promptbench" value="slicer-cache-keep"/></x:ext></x:extLst></x:slicerCache>`);
  zip.file(fixture.comboChartPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:fixture="urn:office-kit:promptbench"><c:chart><c:autoTitleDeleted val="0"/><c:title><c:tx><c:rich><a:t xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${fixture.comboChartMarker}</a:t></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="col"/></c:barChart><c:lineChart><c:grouping val="standard"/></c:lineChart></c:plotArea><c:plotVisOnly val="1"/></c:chart><fixture:preservation value="combo-chart-opaque"/></c:chartSpace>`);
  zip.file("xl/metadata.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:metadata xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:metadataTypes count="1"><x:metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></x:metadataTypes><x:futureMetadata name="XLDAPR" count="1"><x:bk><x:extLst><x:ext uri="{BDBB8CDC-FA1E-496E-A857-3C3F30C029C3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"/></x:ext></x:extLst></x:bk></x:futureMetadata><x:cellMetadata count="1"><x:bk><x:rc t="1" v="0"/></x:bk></x:cellMetadata></x:metadata>`);
  const dataSheetXml = await zip.file(dataSheetPath).async("text");
  if (!/<[^>]*c\b[^>]*\br="G2"[^>]*>/i.test(dataSheetXml)) throw new Error("XLSX opaque-enterprise fixture dynamic-array anchor G2 is missing.");
  return canonicalizeXlsxZip(zip);
}

export async function generateXlsxConnectionRefresh(target) {
  const fixture = XLSX_CONNECTION_REFRESH_FIXTURE;
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add(fixture.sheetName);
  sheet.getRange("A1:B3").values = [
    ["Region", "Revenue"],
    ["North", 120],
    ["South", 90],
  ];
  sheet.tables.add("A1:B3", true, fixture.tableName);
  sheet.getRange("A1:B1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
  sheet.getRange("A1:B3").format.columnWidthPx = 150;
  sheet.freezePanes.freezeRows(1);
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  const patched = await attachConnectionRefreshFixture(exported, fixture);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, patched);
  return { path: target, type: XLSX_MIME };
}

export async function generateXlsxOpaqueEnterprise(target) {
  const fixture = XLSX_OPAQUE_ENTERPRISE_FIXTURE;
  const workbook = Workbook.create();
  const assumptions = workbook.worksheets.add(fixture.assumptionsSheetName);
  assumptions.getRange("A1:C6").values = [
    ["Enterprise planning assumptions", null, null],
    ["Assumption", "Value", "Owner"],
    ["Fiscal year", "FY27", "Finance"],
    ["Revenue growth", fixture.originalAssumption, "Planning"],
    ["Gross margin", 0.58, "Planning"],
    ["Refresh policy", "Manual review", "Data steward"],
  ];
  assumptions.getRange("A1:C1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF", size: 14 } };
  assumptions.getRange("A2:C2").format = { fill: "#E2E8F0", font: { bold: true } };
  assumptions.getRange("B4:B5").setNumberFormat("0.0%");
  assumptions.getRange("A1:C6").format.columnWidthPx = 160;
  assumptions.getRange("A1:A6").format.columnWidthPx = 220;
  assumptions.freezePanes.freezeRows(2);

  const dashboard = workbook.worksheets.add(fixture.dashboardSheetName);
  dashboard.getRange("A1:B4").values = [
    ["Quarter", "Revenue"],
    ["Q1", 100],
    ["Q2", 120],
    ["Q3", 150],
  ];
  dashboard.getRange("A1:B1").format = { fill: "#DBEAFE", font: { bold: true, color: "#1E3A8A" } };
  dashboard.getRange("A1:B4").format.columnWidthPx = 130;
  const chart = dashboard.charts.add("line", dashboard.getRange("A1:B4"));
  chart.name = fixture.chartName;
  chart.title = fixture.originalChartTitle;
  chart.setPosition("D1", "K16");

  const data = workbook.worksheets.add(fixture.dataSheetName);
  data.getRange("A1:D5").values = [
    ["Quarter", "Revenue", "Cost", "Region"],
    ["Q1", 100, 60, "North"],
    ["Q2", 120, 68, "South"],
    ["Q3", 150, 79, "North"],
    ["Q4", 175, 92, "West"],
  ];
  data.tables.add("A1:D5", true, fixture.tableName);
  data.getRange("A1:D1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
  data.getRange("A1:D5").format.columnWidthPx = 120;
  data.freezePanes.freezeRows(1);
  data.getRange("E2:E5").sparklines.add("line", data.getRange("B2:C5"), { seriesColor: "#0EA5E9" });
  const dynamicAnchor = data.store.get("G2");
  // The native metadata part marks this as an XLDAPR-style spill canary, but
  // the bounded public evaluator keeps the legacy array formula result
  // deterministic so workbook.verify() has no formula-error dependency.
  dynamicAnchor.formula = "=1";
  dynamicAnchor.formulaType = "array";
  dynamicAnchor.arrayRef = fixture.dynamicArrayAddress;
  dynamicAnchor.value = 1;
  data.store.get("G3").value = 2;
  data.store.get("G4").value = 3;

  const summary = workbook.worksheets.add(fixture.pivotSheetName);
  summary.getRange("A1:B4").values = [["Enterprise summary", null], ["Region", "Revenue"], ["North", 250], ["South", 120]];
  summary.getRange("A1:B1").format = { fill: "#FEF3C7", font: { bold: true, color: "#92400E" } };
  summary.getRange("A1:B4").format.columnWidthPx = 150;
  summary.pivotTables.add({
    name: fixture.pivotName,
    sourceRange: `${fixture.dataSheetName}!A1:D5`,
    targetRange: "A6",
    rowFields: ["Region"],
    valueFields: [{ field: "Revenue", summarizeBy: "sum", name: "Revenue" }],
    rowGrandTotals: true,
    columnGrandTotals: true,
    refreshPolicy: { refreshOnLoad: true, saveData: true, enableRefresh: true },
  });

  workbook.comments.setSelf({ displayName: "Enterprise planning" });
  const thread = workbook.comments.addThread(
    { cell: assumptions.getRange(fixture.threadedCommentAddress) },
    fixture.threadedCommentText,
    {
      id: "enterprise-assumption-review",
      author: "Planning reviewer",
      resolved: false,
      comment: {
        id: "{44444444-4444-4444-8444-444444444444}",
        personId: "{DDDDDDDD-DDDD-4DDD-8DDD-DDDDDDDDDDDD}",
        date: "2026-07-20T09:00:00.000Z",
        person: {
          id: "{DDDDDDDD-DDDD-4DDD-8DDD-DDDDDDDDDDDD}",
          displayName: "Planning reviewer",
          userId: "planning@example.test",
          providerId: "None",
        },
      },
    },
  );
  thread.addReply("The warehouse refresh remains manual.", {
    id: "{55555555-5555-4555-8555-555555555555}",
    personId: "{EEEEEEEE-EEEE-4EEE-8EEE-EEEEEEEEEEEE}",
    author: "Data steward",
    date: "2026-07-20T09:05:00.000Z",
    person: {
      id: "{EEEEEEEE-EEEE-4EEE-8EEE-EEEEEEEEEEEE}",
      displayName: "Data steward",
      userId: "data@example.test",
      providerId: "None",
    },
    done: true,
  });
  workbook.recalculate();
  const exported = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const patched = await attachOpaqueEnterpriseFixture(exported, fixture);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, patched);
  return { path: target, type: XLSX_MIME };
}

export async function generateXlsxPivotRefresh(target) {
  const fixture = XLSX_PIVOT_REFRESH_FIXTURE;
  const workbook = Workbook.create();
  const data = workbook.worksheets.add(fixture.sourceSheetName);
  data.getRange("A1:B5").values = [
    ["Region", "Revenue"],
    ["East", 120],
    ["West", 90],
    ["East", 30],
    ["North", 60],
  ];
  data.getRange("A1:B1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
  data.getRange("A1:B5").format.columnWidthPx = 128;
  data.freezePanes.freezeRows(1);
  const summary = workbook.worksheets.add(fixture.sheetName);
  summary.getRange("A1:B5").format = { border: { bottom: { style: "thin", color: "#CBD5E1" } } };
  summary.getRange("A1:B1").format = { fill: "#DBEAFE", font: { bold: true, color: "#1E3A8A" } };
  summary.getRange("A1:B5").format.columnWidthPx = 144;
  summary.pivotTables.add({
    name: fixture.pivotName,
    sourceRange: fixture.sourceRange,
    targetRange: fixture.targetRange,
    rowFields: ["Region"],
    valueFields: [{ field: "Revenue", summarizeBy: "sum", name: "Revenue" }],
    rowGrandTotals: true,
    columnGrandTotals: true,
    refreshPolicy: { refreshOnLoad: true, saveData: true, enableRefresh: true },
  });
  const exported = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const bytes = new Uint8Array(await exported.arrayBuffer());
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
  const cacheDefinitions = paths.filter((name) => /^pivotCache\/pivotCacheDefinition\d*\.xml$/i.test(name));
  const pivotTables = paths.filter((name) => /^xl\/pivotTables\/pivotTable\d*\.xml$/i.test(name));
  if (cacheDefinitions.length !== 1 || pivotTables.length !== 1) {
    throw new Error("XLSX Pivot refresh fixture must contain exactly one native PivotTable and cache definition.");
  }
  const cacheXml = await zip.file(cacheDefinitions[0]).async("text");
  if (!/refreshOnLoad="(?:1|true)"/i.test(cacheXml)) {
    throw new Error("XLSX Pivot refresh fixture did not author an explicit cache refreshOnLoad=true request.");
  }
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, bytes);
  return { path: target, type: XLSX_MIME };
}

export async function generateDocxClassicCommentReview(target) {
  const fixture = DOCX_CLASSIC_COMMENT_FIXTURE;
  const document = DocumentModel.create({
    name: fixture.title,
    defaultRunStyle: { fontFamily: "Aptos", fontSize: 11, color: "#172033" },
    blocks: [],
  });
  document.addParagraph(fixture.title, {
    paragraphFormat: { spaceAfterTwips: 160 },
    runs: [{ text: fixture.title, style: { bold: true, fontSize: 16, color: "#123B5D" } }],
  });
  const decision = document.addParagraph(fixture.anchorText, {
    paragraphFormat: { spaceAfterTwips: 120 },
    runs: [{ text: fixture.anchorText, style: { bold: true } }],
  });
  document.addParagraph(fixture.supportingText, {
    paragraphFormat: { spaceAfterTwips: 120 },
  });
  document.addParagraph("Reviewer instruction: preserve the decision text and update only the attached classic comment.");
  document.addComment(decision, fixture.comment.originalText, {
    author: fixture.comment.author,
    initials: fixture.comment.initials,
    date: fixture.comment.date,
  });
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated DOCX classic-comment fixture failed model verification: " + verification.ndjson);
  const exported = await DocumentFile.exportDocx(document);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await exported.arrayBuffer()));
  return { path: target, type: DOCX_MIME };
}

export async function generateDocxModernCommentReplyBoundary(target) {
  const fixture = DOCX_MODERN_COMMENT_REPLY_BOUNDARY_FIXTURE;
  const document = DocumentModel.create({ name: fixture.title, blocks: [] });
  const anchor = document.addParagraph(fixture.anchorText);
  const root = document.addComment(anchor, fixture.root.text, {
    author: fixture.root.author,
    initials: fixture.root.initials,
    date: fixture.root.date,
    resolved: false,
    paraId: fixture.root.paraId,
    durableId: fixture.root.durableId,
    dateUtc: fixture.root.date,
    person: { providerId: fixture.root.providerId, userId: fixture.root.userId },
  });
  document.replyToComment(root, fixture.directReply.text, {
    author: fixture.directReply.author,
    initials: fixture.directReply.initials,
    date: fixture.directReply.date,
    resolved: false,
    paraId: fixture.directReply.paraId,
    durableId: fixture.directReply.durableId,
    dateUtc: fixture.directReply.date,
    person: { providerId: fixture.directReply.providerId, userId: fixture.directReply.userId },
  });
  const exported = await DocumentFile.exportDocx(document);
  const zip = await JSZip.loadAsync(await exported.arrayBuffer());
  const nested = fixture.nestedReply;
  const commentXml = await zip.file("word/comments.xml")?.async("text");
  const commentsExtendedXml = await zip.file("word/commentsExtended.xml")?.async("text");
  const commentsIdsXml = await zip.file("word/commentsIds.xml")?.async("text");
  const peopleXml = await zip.file("word/people.xml")?.async("text");
  if (!commentXml || !commentsExtendedXml || !commentsIdsXml || !peopleXml) {
    throw new Error("modern comment fixture is missing a canonical identity part");
  }
  zip.file("word/comments.xml", commentXml.replace(
    "</w:comments>",
    `<w:comment w:initials="${xmlEscape(nested.initials)}" w:author="${xmlEscape(nested.author)}" w:date="${nested.date}" w:id="${nested.id}"><w:p w14:paraId="${nested.paraId}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:r><w:t>${xmlEscape(nested.text)}</w:t></w:r></w:p></w:comment></w:comments>`,
  ));
  zip.file("word/commentsExtended.xml", commentsExtendedXml.replace(
    "</w15:commentsEx>",
    `<w15:commentEx w15:paraId="${nested.paraId}" w15:paraIdParent="${nested.parentParaId}" w15:done="0" /></w15:commentsEx>`,
  ));
  zip.file("word/commentsIds.xml", commentsIdsXml.replace(
    "</w16cid:commentsIds>",
    `<w16cid:commentId w16cid:paraId="${nested.paraId}" w16cid:durableId="${nested.durableId}" /></w16cid:commentsIds>`,
  ));
  zip.file("word/people.xml", peopleXml.replace(
    "</w15:people>",
    `<w15:person w15:author="${xmlEscape(nested.author)}"><w15:presenceInfo w15:providerId="${xmlEscape(nested.providerId)}" w15:userId="${xmlEscape(nested.userId)}" /></w15:person></w15:people>`,
  ));

  // The OfficeKit exporter allocates package-local relationship IDs and ZIP
  // timestamps dynamically.  Canonicalize both so the locked fixture can be
  // regenerated byte-for-byte in every PromptBench trial.
  const entries = [];
  for (const name of Object.keys(zip.files).filter((entry) => !zip.files[entry].dir).sort()) {
    const file = zip.file(name);
    if (!file) continue;
    entries.push({ name, bytes: await file.async("uint8array") });
  }
  const relationshipIds = new Map();
  const xmlEntries = entries.map(({ name, bytes }) => {
    if (!/\.(?:xml|rels)$/i.test(name)) return { name, bytes, xml: null };
    const xml = new TextDecoder().decode(bytes);
    for (const match of xml.matchAll(/\bR[a-f0-9]{16}\b/gi)) {
      if (!relationshipIds.has(match[0])) relationshipIds.set(match[0], `rId${relationshipIds.size + 1}`);
    }
    return { name, bytes, xml };
  });
  const canonicalZip = new JSZip();
  const fixedDate = new Date("1980-01-01T00:00:00.000Z");
  for (const entry of xmlEntries) {
    const data = entry.xml === null
      ? entry.bytes
      : entry.xml.replace(/\bR[a-f0-9]{16}\b/gi, (id) => relationshipIds.get(id) || id);
    canonicalZip.file(entry.name, data, {
      date: fixedDate,
      createFolders: false,
      compression: "DEFLATE",
      compressionOptions: { level: 9 },
    });
  }
  const bytes = await canonicalZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 9 }, platform: "DOS" });
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, bytes);
  return { path: target, type: DOCX_MIME };
}

export async function generateDocxComplexTableTopologyBoundary(target) {
  const fixture = DOCX_COMPLEX_TABLE_TOPOLOGY_FIXTURE;
  const document = DocumentModel.create({ name: fixture.title, blocks: [] });
  document.styles.add("TableGrid", { name: "Table Grid", type: "table" });
  document.addParagraph(fixture.title, { runs: [{ text: fixture.title, style: { bold: true, fontSize: 16 } }] });
  document.addParagraph(fixture.leadText);
  document.addTable({
    name: "ordinary-canary",
    values: [["Canary", "Value"], [fixture.canaryText, "Stable"]],
  });
  document.addParagraph("The following imported table intentionally combines source-owned topology.");
  document.addTable({
    name: "complex-source-bound",
    values: [fixture.complexTable.headers, ["Antibiotic", "500 mg", "PO", "Pending review"]],
  });
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated DOCX complex-table base failed model verification: " + verification.ndjson);
  const exported = await DocumentFile.exportDocx(document);
  const zip = await JSZip.loadAsync(await exported.arrayBuffer());
  const documentXml = await zip.file("word/document.xml")?.async("text");
  const stylesXml = await zip.file("word/styles.xml")?.async("text");
  if (!documentXml || !stylesXml) throw new Error("complex-table fixture is missing document or styles XML");
  const tableMatches = [...documentXml.matchAll(/<w:tbl\b[\s\S]*?<\/w:tbl>/g)];
  if (tableMatches.length !== 2) throw new Error(`complex-table fixture expected two base tables, found ${tableMatches.length}`);

  const run = (text) => `<w:r><w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r>`;
  const paragraph = (body) => `<w:p>${body}</w:p>`;
  const cell = (body, properties = "") => `<w:tc>${properties ? `<w:tcPr>${properties}</w:tcPr>` : ""}${body}</w:tc>`;
  const width = (value) => `<w:tcW w:w="${value}" w:type="dxa"/>`;
  const nestedTable = `<w:tbl><w:tblPr><w:tblW w:w="1800" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1800"/></w:tblGrid><w:tr>${cell(paragraph(run(fixture.complexTable.nestedCell)), width(1800))}</w:tr></w:tbl>`;
  const contentControl = `<w:sdt><w:sdtPr><w:id w:val="9001"/><w:alias w:val="Status"/></w:sdtPr><w:sdtContent>${paragraph(run(fixture.complexTable.contentControl))}</w:sdtContent></w:sdt>`;
  const revised = `<w:ins w:id="41" w:author="Clinical reviewer" w:date="2026-07-19T09:00:00Z">${run(fixture.complexTable.revisedCell)}</w:ins>`;
  const complexTable = [
    `<w:tbl><w:tblPr><w:tblStyle w:val="${fixture.complexTable.styleId}"/><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>`,
    "<w:tblGrid><w:gridCol w:w=\"2200\"/><w:gridCol w:w=\"1400\"/><w:gridCol w:w=\"1400\"/><w:gridCol w:w=\"2200\"/></w:tblGrid>",
    `<w:tr>${fixture.complexTable.headers.map((header) => cell(paragraph(run(header)), width(1800))).join("")}</w:tr>`,
    `<w:tr>${cell(paragraph(run(fixture.complexTable.mergeRoot)), `${width(2200)}<w:vMerge w:val="restart"/>`)}${cell(paragraph(run("500 mg")), width(1400))}${cell(paragraph(run("PO")), width(1400))}${cell(contentControl, width(2200))}</w:tr>`,
    `<w:tr>${cell(paragraph(run(fixture.complexTable.mergeContinuation)), `${width(2200)}<w:vMerge/>`)}${cell(paragraph(revised), width(1400))}${cell(paragraph(run("IV")), width(1400))}${cell(`${paragraph(run("Review details"))}${nestedTable}`, width(2200))}</w:tr>`,
    `<w:tr>${cell(paragraph(run("Dose adjustment")), `${width(2200)}<w:vMerge w:val="restart"/>`)}${cell(paragraph(run("250 mg")), width(1400))}${cell(paragraph(run("IM")), width(1400))}${cell(paragraph(run("Pending review")), width(2200))}</w:tr>`,
    "</w:tbl>",
  ].join("");
  const firstTableEnd = tableMatches[1].index;
  const firstTable = documentXml.slice(0, firstTableEnd);
  const secondTableStart = tableMatches[1].index + tableMatches[1][0].length;
  const patchedDocumentXml = `${firstTable}${complexTable}${documentXml.slice(secondTableStart)}`;
  const patchedStylesXml = stylesXml.includes(`w:styleId="${fixture.complexTable.styleId}"`)
    ? stylesXml
    : stylesXml.replace("</w:styles>", `<w:style w:type="table" w:styleId="${fixture.complexTable.styleId}"><w:name w:val="Clinical Grid"/></w:style></w:styles>`);
  zip.file("word/document.xml", patchedDocumentXml);
  zip.file("word/styles.xml", patchedStylesXml);

  // Canonicalize the package so the checked-in source is reproducible across
  // runs even though the exporter allocates relationship IDs and ZIP dates.
  const entries = [];
  for (const name of Object.keys(zip.files).filter((entry) => !zip.files[entry].dir).sort()) {
    const file = zip.file(name);
    if (file) entries.push({ name, bytes: await file.async("uint8array") });
  }
  const relationshipIds = new Map();
  const xmlEntries = entries.map(({ name, bytes }) => {
    if (!/\.(?:xml|rels)$/i.test(name)) return { name, bytes, xml: null };
    const xml = new TextDecoder().decode(bytes);
    for (const match of xml.matchAll(/\bR[a-f0-9]{16}\b/gi)) {
      if (!relationshipIds.has(match[0])) relationshipIds.set(match[0], `rId${relationshipIds.size + 1}`);
    }
    return { name, bytes, xml };
  });
  const canonicalZip = new JSZip();
  const fixedDate = new Date("1980-01-01T00:00:00.000Z");
  for (const entry of xmlEntries) {
    const data = entry.xml === null
      ? entry.bytes
      : entry.xml.replace(/\bR[a-f0-9]{16}\b/gi, (id) => relationshipIds.get(id) || id);
    canonicalZip.file(entry.name, data, {
      date: fixedDate,
      createFolders: false,
      compression: "DEFLATE",
      compressionOptions: { level: 9 },
    });
  }
  const bytes = await canonicalZip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
    platform: "DOS",
  });
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, bytes);
  return { path: target, type: DOCX_MIME };
}

async function canonicalizeDocxFixtureZip(zip) {
  const entries = [];
  for (const name of Object.keys(zip.files).filter((entry) => !zip.files[entry].dir).sort()) {
    const file = zip.file(name);
    if (file) entries.push({ name, bytes: await file.async("uint8array") });
  }
  const relationshipIds = new Map();
  const xmlEntries = entries.map(({ name, bytes }) => {
    if (!/\.(?:xml|rels)$/i.test(name)) return { name, bytes, xml: null };
    let xml = new TextDecoder().decode(bytes);
    // OfficeKit intentionally allocates fresh classic-comment paragraph and
    // durable IDs. The fixture contract binds those identities, so map the
    // generated classic root to fixed values before publishing the asset.
    if (name === "word/comments.xml") {
      xml = xml.replace(/(w:id="0">[\s\S]*?w14:paraId=")[A-F0-9]{8}(")/i, (_match, prefix, suffix) => `${prefix}758DCE26${suffix}`);
    } else if (name === "word/commentsExtended.xml") {
      xml = xml.replace(/(w15:commentEx w15:paraId=")[A-F0-9]{8}(")/i, (_match, prefix, suffix) => `${prefix}758DCE26${suffix}`);
    } else if (name === "word/commentsIds.xml") {
      xml = xml.replace(/(w16cid:commentId w16cid:paraId=")[A-F0-9]{8}(")/i, (_match, prefix, suffix) => `${prefix}758DCE26${suffix}`);
      xml = xml.replace(/(w16cid:commentId w16cid:paraId="758DCE26" w16cid:durableId=")[A-F0-9]{8}(")/i, (_match, prefix, suffix) => `${prefix}6A47B40A${suffix}`);
    } else if (name === "word/commentsExtensible.xml") {
      xml = xml.replace(/(w16cex:commentExtensible w16cex:durableId=")[A-F0-9]{8}(")/i, (_match, prefix, suffix) => `${prefix}6A47B40A${suffix}`);
    }
    for (const match of xml.matchAll(/\bR[a-f0-9]{16}\b/gi)) {
      if (!relationshipIds.has(match[0])) relationshipIds.set(match[0], `rId${relationshipIds.size + 1}`);
    }
    return { name, bytes, xml };
  });
  const canonicalZip = new JSZip();
  const fixedDate = new Date("1980-01-01T00:00:00.000Z");
  for (const entry of xmlEntries) {
    const data = entry.xml === null
      ? entry.bytes
      : entry.xml.replace(/\bR[a-f0-9]{16}\b/gi, (id) => relationshipIds.get(id) || id);
    canonicalZip.file(entry.name, data, {
      date: fixedDate,
      createFolders: false,
      compression: "DEFLATE",
      compressionOptions: { level: 9 },
    });
  }
  return canonicalZip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
    platform: "DOS",
  });
}

export async function generateDocxSurgicalBoardReview(target) {
  const fixture = DOCX_SURGICAL_BOARD_REVIEW_FIXTURE;
  const document = DocumentModel.create({
    name: fixture.title,
    blocks: [],
    defaultRunStyle: { fontFamily: "Aptos", fontSize: 11, color: "#172033" },
  });
  document.styles.add("TableGrid", { name: "Table Grid", type: "table" });
  document.addParagraph(fixture.title, {
    paragraphFormat: { spaceAfterTwips: 160 },
    runs: [{ text: fixture.title, style: { bold: true, fontSize: 16, color: "#123B5D" } }],
  });
  const recommendation = document.addParagraph(fixture.recommendation.originalText, {
    paragraphFormat: { spaceAfterTwips: 120 },
    runs: [{ text: fixture.recommendation.originalText, style: { bold: true } }],
  });
  document.addBookmark(recommendation, fixture.refBookmark, { nativeId: 41 });
  document.addComment(recommendation, fixture.comment.originalText, {
    author: fixture.comment.author,
    initials: fixture.comment.initials,
    date: fixture.comment.date,
  });
  document.addFootnote(recommendation, fixture.footnoteText, {
    id: 7,
    author: "Board secretary",
  });
  document.addTable({
    name: "risk-matrix",
    values: [
      fixture.riskTable.headers,
      [fixture.riskTable.targetLabel, fixture.riskTable.originalStatus],
      ["Security controls", "Green"],
    ],
  });
  document.addParagraph("The following table and review records remain source-owned.");
  document.addTable({
    name: "complex-source-bound",
    values: [fixture.complexTable.headers, ["Antibiotic", "500 mg", "PO", "Pending review"]],
  });
  document.addTableOfContents({ levels: "1-3", display: "Board contents" });
  const owner = document.addParagraph("Review owner");
  owner.addTextContentControl("Audit committee", { tag: "review-owner", alias: "Review owner" });
  document.addInsertion("Pending revision remains in the source.", {
    author: "Clinical reviewer",
    date: "2026-07-19T09:00:00Z",
    id: 41,
  });
  const modernAnchor = document.addParagraph("Modern review thread remains source-owned.");
  const root = document.addComment(modernAnchor, "Please confirm the release evidence.", {
    author: "Lead reviewer",
    initials: "LR",
    date: "2026-07-19T08:00:00Z",
    resolved: false,
    paraId: fixture.modernComment.rootParaId,
    durableId: "17777770",
    dateUtc: "2026-07-19T08:00:00Z",
    person: { providerId: "provider-a", userId: "lead@example.test" },
  });
  document.replyToComment(root, "The evidence is attached.", {
    author: "Release reviewer",
    initials: "RR",
    date: "2026-07-19T08:05:00Z",
    resolved: false,
    paraId: fixture.modernComment.directReplyParaId,
    durableId: "18888880",
    dateUtc: "2026-07-19T08:05:00Z",
    person: { providerId: "provider-b", userId: "release@example.test" },
  });
  document.addSection({ breakType: "nextPage", pageNumbering: { start: 1, format: "lowerRoman" } });
  document.addParagraph("Front matter canary — preserve the review packet.");
  document.addSection({ breakType: "nextPage", pageNumbering: { start: 1, format: "decimal" } });
  document.addParagraph("Body section canary — preserve pagination and fields.");
  for (let sectionIndex = 0; sectionIndex < fixture.sectionCount; sectionIndex += 1) {
    document.addHeader(fixture.headerTexts[0], { sectionIndex, referenceType: "default" });
    document.addHeader(fixture.headerTexts[1], { sectionIndex, referenceType: "first" });
    document.addHeader(fixture.headerTexts[2], { sectionIndex, referenceType: "even" });
    document.addFooter("1", { sectionIndex, referenceType: "default", fieldInstruction: "PAGE" });
  }
  document.addWatermark(fixture.watermarkText, { id: "watermark/board-review" });
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated DOCX board-review base failed model verification: " + verification.ndjson);

  const exported = await DocumentFile.exportDocx(document);
  const zip = await JSZip.loadAsync(await exported.arrayBuffer());
  let documentXml = await zip.file("word/document.xml")?.async("text");
  const stylesXml = await zip.file("word/styles.xml")?.async("text");
  if (!documentXml || !stylesXml) throw new Error("board-review fixture is missing document or styles XML");
  const tableMatches = [...documentXml.matchAll(/<w:tbl\b[\s\S]*?<\/w:tbl>/g)];
  if (tableMatches.length !== 2) throw new Error(`board-review fixture expected two base tables, found ${tableMatches.length}`);

  const run = (text) => `<w:r><w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r>`;
  const paragraph = (body) => `<w:p>${body}</w:p>`;
  const cell = (body, properties = "") => `<w:tc>${properties ? `<w:tcPr>${properties}</w:tcPr>` : ""}${body}</w:tc>`;
  const width = (value) => `<w:tcW w:w="${value}" w:type="dxa"/>`;
  const nestedTable = `<w:tbl><w:tblPr><w:tblW w:w="1800" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1800"/></w:tblGrid><w:tr>${cell(paragraph(run(fixture.complexTable.nestedCell)), width(1800))}</w:tr></w:tbl>`;
  const contentControl = `<w:sdt><w:sdtPr><w:id w:val="9001"/><w:alias w:val="Status"/></w:sdtPr><w:sdtContent>${paragraph(run(fixture.complexTable.contentControl))}</w:sdtContent></w:sdt>`;
  const revised = `<w:ins w:id="41" w:author="Clinical reviewer" w:date="2026-07-19T09:00:00Z">${run(fixture.complexTable.revisedCell)}</w:ins>`;
  const complexTable = [
    `<w:tbl><w:tblPr><w:tblStyle w:val="${fixture.complexTable.styleId}"/><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>`,
    "<w:tblGrid><w:gridCol w:w=\"2200\"/><w:gridCol w:w=\"1400\"/><w:gridCol w:w=\"1400\"/><w:gridCol w:w=\"2200\"/></w:tblGrid>",
    `<w:tr>${fixture.complexTable.headers.map((header) => cell(paragraph(run(header)), width(1800))).join("")}</w:tr>`,
    `<w:tr>${cell(paragraph(run(fixture.complexTable.mergeRoot)), `${width(2200)}<w:vMerge w:val="restart"/>`)}${cell(paragraph(run("500 mg")), width(1400))}${cell(paragraph(run("PO")), width(1400))}${cell(contentControl, width(2200))}</w:tr>`,
    `<w:tr>${cell(paragraph(run(fixture.complexTable.mergeContinuation)), `${width(2200)}<w:vMerge/>`)}${cell(paragraph(revised), width(1400))}${cell(paragraph(run("IV")), width(1400))}${cell(`${paragraph(run("Review details"))}${nestedTable}`, width(2200))}</w:tr>`,
    `<w:tr>${cell(paragraph(run("Dose adjustment")), `${width(2200)}<w:vMerge w:val="restart"/>`)}${cell(paragraph(run("250 mg")), width(1400))}${cell(paragraph(run("IM")), width(1400))}${cell(paragraph(run("Pending review")), width(2200))}</w:tr>`,
    "</w:tbl>",
  ].join("");
  const firstTableEnd = tableMatches[1].index;
  const firstTable = documentXml.slice(0, firstTableEnd);
  const secondTableStart = tableMatches[1].index + tableMatches[1][0].length;
  documentXml = `${firstTable}${complexTable}${documentXml.slice(secondTableStart)}`;
  const refField = `<w:p><w:fldSimple w:instr=" REF ${fixture.refBookmark} "><w:r><w:t>${xmlEscape(fixture.recommendation.originalText)}</w:t></w:r></w:fldSimple></w:p>`;
  const finalSection = documentXml.lastIndexOf("<w:sectPr");
  if (finalSection < 0) throw new Error("board-review fixture has no final section properties");
  documentXml = documentXml.slice(0, finalSection) + refField + documentXml.slice(finalSection);
  zip.file("word/document.xml", documentXml);
  zip.file("word/styles.xml", stylesXml.includes(`w:styleId="${fixture.complexTable.styleId}"`)
    ? stylesXml
    : stylesXml.replace("</w:styles>", `<w:style w:type="table" w:styleId="${fixture.complexTable.styleId}"><w:name w:val="Board Review Grid"/></w:style></w:styles>`));

  const commentXml = await zip.file("word/comments.xml")?.async("text");
  const commentsExtendedXml = await zip.file("word/commentsExtended.xml")?.async("text");
  const commentsIdsXml = await zip.file("word/commentsIds.xml")?.async("text");
  const peopleXml = await zip.file("word/people.xml")?.async("text");
  if (!commentXml || !commentsExtendedXml || !commentsIdsXml || !peopleXml) throw new Error("board-review fixture is missing modern comment identity parts");
  zip.file("word/comments.xml", commentXml.replace(
    "</w:comments>",
    `<w:comment w:initials="NR" w:author="Nested reviewer" w:date="2026-07-19T08:10:00Z" w:id="3"><w:p w14:paraId="${fixture.modernComment.nestedReplyParaId}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:r><w:t>Approved after legal review.</w:t></w:r></w:p></w:comment></w:comments>`,
  ));
  zip.file("word/commentsExtended.xml", commentsExtendedXml.replace(
    "</w15:commentsEx>",
    `<w15:commentEx w15:paraId="${fixture.modernComment.nestedReplyParaId}" w15:paraIdParent="${fixture.modernComment.directReplyParaId}" w15:done="0" /></w15:commentsEx>`,
  ));
  zip.file("word/commentsIds.xml", commentsIdsXml.replace(
    "</w16cid:commentsIds>",
    `<w16cid:commentId w16cid:paraId="${fixture.modernComment.nestedReplyParaId}" w16cid:durableId="19999990" /></w16cid:commentsIds>`,
  ));
  zip.file("word/people.xml", peopleXml.replace(
    "</w15:people>",
    `<w15:person w15:author="Nested reviewer"><w15:presenceInfo w15:providerId="provider-c" w15:userId="nested@example.test" /></w15:person></w15:people>`,
  ));
  const bytes = await canonicalizeDocxFixtureZip(zip);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, bytes);
  return { path: target, type: DOCX_MIME };
}

export async function generateDocxHeaderTextReview(target) {
  const fixture = DOCX_HEADER_TEXT_FIXTURE;
  const document = DocumentModel.create({
    name: fixture.title,
    defaultRunStyle: { fontFamily: "Aptos", fontSize: 11, color: "#172033" },
    blocks: [],
  });
  document.addParagraph(fixture.title, {
    paragraphFormat: { spaceAfterTwips: 160 },
    runs: [{ text: fixture.title, style: { bold: true, fontSize: 16, color: "#123B5D" } }],
  });
  for (const text of fixture.body) document.addParagraph(text, { paragraphFormat: { spaceAfterTwips: 120 } });
  document.addHeader(fixture.header.originalText, {
    id: "header/review-target",
    sectionIndex: fixture.header.sectionIndex,
    referenceType: fixture.header.referenceType,
  });
  document.addHeader(fixture.header.companionText, {
    id: "header/companion",
    sectionIndex: fixture.header.sectionIndex,
    referenceType: fixture.header.referenceType,
  });
  document.addFooter(fixture.footer.text, {
    id: "footer/page",
    sectionIndex: fixture.header.sectionIndex,
    referenceType: fixture.header.referenceType,
    fieldInstruction: fixture.footer.fieldInstruction,
  });
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated DOCX header-text fixture failed model verification: " + verification.ndjson);
  const exported = await DocumentFile.exportDocx(document);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await exported.arrayBuffer()));
  return { path: target, type: DOCX_MIME };
}

export async function generateDocxFooterTextReview(target) {
  const fixture = DOCX_FOOTER_TEXT_FIXTURE;
  const document = DocumentModel.create({
    name: fixture.title,
    defaultRunStyle: { fontFamily: "Aptos", fontSize: 11, color: "#172033" },
    blocks: [],
  });
  document.addParagraph(fixture.title, {
    paragraphFormat: { spaceAfterTwips: 160 },
    runs: [{ text: fixture.title, style: { bold: true, fontSize: 16, color: "#123B5D" } }],
  });
  for (const text of fixture.body) document.addParagraph(text, { paragraphFormat: { spaceAfterTwips: 120 } });
  document.addHeader(fixture.header.text, {
    id: "header/page",
    sectionIndex: fixture.footer.sectionIndex,
    referenceType: fixture.footer.referenceType,
    fieldInstruction: fixture.header.fieldInstruction,
  });
  document.addFooter(fixture.footer.originalText, {
    id: "footer/review-target",
    sectionIndex: fixture.footer.sectionIndex,
    referenceType: fixture.footer.referenceType,
  });
  document.addFooter(fixture.footer.companionText, {
    id: "footer/companion",
    sectionIndex: fixture.footer.sectionIndex,
    referenceType: fixture.footer.referenceType,
  });
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated DOCX footer-text fixture failed model verification: " + verification.ndjson);
  const exported = await DocumentFile.exportDocx(document);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await exported.arrayBuffer()));
  return { path: target, type: DOCX_MIME };
}

export async function generateDocxSectionPageNumberingReview(target) {
  const fixture = DOCX_SECTION_PAGE_NUMBERING_FIXTURE;
  const document = DocumentModel.create({
    name: fixture.title,
    defaultRunStyle: { fontFamily: "Aptos", fontSize: 11, color: "#172033" },
    blocks: [],
  });
  document.addParagraph(fixture.title, {
    paragraphFormat: { spaceAfterTwips: 160 },
    runs: [{ text: fixture.title, style: { bold: true, fontSize: 16, color: "#123B5D" } }],
  });
  document.addSection({ breakType: "nextPage", pageNumbering: fixture.target.originalPageNumbering });
  document.addParagraph(fixture.body[0], { paragraphFormat: { spaceAfterTwips: 120 } });
  document.addSection({ breakType: "nextPage", pageNumbering: fixture.sibling.pageNumbering });
  document.addParagraph(fixture.body[1], { paragraphFormat: { spaceAfterTwips: 120 } });
  document.addParagraph(fixture.body[2], { paragraphFormat: { spaceAfterTwips: 120 } });
  for (let sectionIndex = 0; sectionIndex < fixture.footerSectionCount; sectionIndex += 1) {
    document.addFooter(fixture.footer.text, {
      id: `footer/page-${sectionIndex}`,
      sectionIndex,
      referenceType: "default",
      fieldInstruction: fixture.footer.fieldInstruction,
    });
  }
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated DOCX section page-numbering fixture failed model verification: " + verification.ndjson);
  const exported = await DocumentFile.exportDocx(document);
  const bytes = new Uint8Array(await exported.arrayBuffer());
  const imported = await DocumentFile.importDocx(bytes);
  const targetSection = imported.blocks[fixture.target.blockIndex];
  const siblingSection = imported.blocks[fixture.sibling.blockIndex];
  if (targetSection?.kind !== "section" || !targetSection.editable
    || JSON.stringify(targetSection.pageNumbering) !== JSON.stringify(fixture.target.originalPageNumbering)) {
    throw new Error("Generated DOCX page-numbering fixture did not reimport its canonical target section.");
  }
  if (siblingSection?.kind !== "section" || !siblingSection.editable
    || JSON.stringify(siblingSection.pageNumbering) !== JSON.stringify(fixture.sibling.pageNumbering)) {
    throw new Error("Generated DOCX page-numbering fixture did not reimport its canonical sibling section.");
  }
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, bytes);
  return { path: target, type: DOCX_MIME };
}


export async function generateOfficeInput(generator, target) {
  if (generator === "xlsx-threaded-review") return generateXlsxThreadedReview(target);
  if (generator === "xlsx-nested-reply-boundary") return generateXlsxNestedReplyBoundary(target);
  if (generator === "xlsx-growth-update") return generateXlsxGrowthUpdate(target);
  if (generator === "xlsx-connection-refresh") return generateXlsxConnectionRefresh(target);
  if (generator === "xlsx-opaque-enterprise") return generateXlsxOpaqueEnterprise(target);
  if (generator === "xlsx-pivot-refresh") return generateXlsxPivotRefresh(target);
  if (generator === "docx-classic-comment-review") return generateDocxClassicCommentReview(target);
  if (generator === "docx-modern-comment-reply-boundary") return generateDocxModernCommentReplyBoundary(target);
  if (generator === "docx-complex-table-topology-boundary") return generateDocxComplexTableTopologyBoundary(target);
  if (generator === "docx-surgical-board-review") return generateDocxSurgicalBoardReview(target);
  if (generator === "docx-header-text-review") return generateDocxHeaderTextReview(target);
  if (generator === "docx-footer-text-review") return generateDocxFooterTextReview(target);
  if (generator === "docx-section-page-numbering-review") return generateDocxSectionPageNumberingReview(target);
  return null;
}
