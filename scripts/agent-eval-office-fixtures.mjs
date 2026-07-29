import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

import {
  DocumentFile,
  DocumentModel,
  Presentation,
  PresentationFile,
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
    providerId: "None",
    date: "2026-07-17T09:00:00.000Z",
    text: "Please confirm the downside cash buffer before board circulation.",
  }),
  priorReply: Object.freeze({
    id: "{22222222-2222-4222-8222-222222222222}",
    personId: "{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}",
    author: "Risk reviewer",
    userId: "risk.reviewer@example.com",
    providerId: "None",
    date: "2026-07-17T09:30:00.000Z",
    text: "Sensitivity analysis is attached to the approved planning case.",
  }),
  requestedReply: "Approved after sensitivity review",
});

// The source-backed threaded-comment profile is deliberately limited to one
// root plus direct replies. This fixture starts from that accepted native XLSX
// package and appends a real reply-of-reply with its own person identity. It
// is therefore an Office-valid relationship/content-type graph that must be
// retained and refused, rather than a malformed parent-reference control.
export const XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE = Object.freeze({
  workbookName: "reviewed-budget-nested.xlsx",
  sheetName: XLSX_THREADED_REVIEW_FIXTURE.sheetName,
  address: XLSX_THREADED_REVIEW_FIXTURE.address,
  worksheetPartPath: "xl/worksheets/sheet1.xml",
  threadedPartPath: "xl/threadedcomments/threadedcomment.xml",
  personPartPath: "xl/persons/person.xml",
  root: XLSX_THREADED_REVIEW_FIXTURE.root,
  directReply: XLSX_THREADED_REVIEW_FIXTURE.priorReply,
  nestedReply: Object.freeze({
    id: "{33333333-3333-4333-8333-333333333333}",
    personId: "{CCCCCCCC-CCCC-4CCC-8CCC-CCCCCCCCCCCC}",
    author: "Legal reviewer",
    userId: "legal.reviewer@example.com",
    providerId: "None",
    date: "2026-07-17T09:45:00.000Z",
    text: "Legal review is in progress.",
    done: false,
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

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

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

// The public document model intentionally supports only a root plus direct
// replies. This fixture starts with that supported profile and then adds a
// third, genuinely nested native Word reply. It is therefore a real Office
// package that must be preserved and refused, not a malformed self-parenting
// graph invented solely for the evaluator.
export const DOCX_MODERN_COMMENT_REPLY_BOUNDARY_FIXTURE = Object.freeze({
  documentName: "modern-comment-replies.docx",
  title: "Controlled rollout modern review",
  anchorText: "Decision: preserve the source-bound modern review thread.",
  supportingText: "The nested legal reply is an immutable review-history canary.",
  comments: Object.freeze([
    Object.freeze({
      nativeId: "0",
      paraId: "11111111",
      durableId: "33333333",
      author: "Lead reviewer",
      initials: "LR",
      date: "2026-07-20T09:00:00Z",
      dateUtc: "2026-07-20T09:00:00Z",
      providerId: "provider-a",
      userId: "lead@example.test",
      text: "Please confirm the legal boundary.",
      resolved: true,
      parentNativeId: "",
    }),
    Object.freeze({
      nativeId: "1",
      paraId: "22222222",
      durableId: "44444444",
      author: "Release reviewer",
      initials: "RR",
      date: "2026-07-20T09:05:00Z",
      dateUtc: "2026-07-20T09:05:00Z",
      providerId: "provider-b",
      userId: "release@example.test",
      text: "Evidence attached for release review.",
      resolved: false,
      parentNativeId: "0",
    }),
    Object.freeze({
      nativeId: "2",
      paraId: "55555555",
      durableId: "66666666",
      author: "Legal reviewer",
      initials: "LG",
      date: "2026-07-20T09:10:00Z",
      dateUtc: "2026-07-20T09:10:00Z",
      providerId: "provider-c",
      userId: "legal@example.test",
      text: "Legal review is in progress.",
      resolved: false,
      parentNativeId: "1",
    }),
  ]),
});

// This is a deliberately small imported-table boundary, not a claim that
// OfficeKit can author or reconstruct arbitrary Word tables. The second
// top-level table combines exactly one vertical merge, nested table, custom
// table style, revision, and block-level SDT. Adding a column through that
// profile would have to preserve several independent native graphs, so the
// public model must inspect and refuse rather than approximate the edit.
export const DOCX_COMPLEX_TABLE_TOPOLOGY_BOUNDARY_FIXTURE = Object.freeze({
  documentName: "clinical-form.docx",
  title: "Clinical protocol review",
  introduction: "The following source-bound clinical dose matrix is retained for review only.",
  baselineTable: Object.freeze([
    Object.freeze(["Record", "Status"]),
    Object.freeze(["Protocol intake", "Complete"]),
  ]),
  tableCaption: "ClinicalDoseMatrix",
  styleId: "ClinicalProtocolTable",
  headers: Object.freeze(["Dose", "Route", "Review state"]),
  mergedDose: "5 mg",
  routeValue: "Oral",
  nestedScheduleLabel: "Schedule details",
  nestedSchedule: Object.freeze(["Morning", "With food"]),
  revision: Object.freeze({
    id: "37",
    author: "Clinical QA",
    date: "2026-07-28T08:00:00Z",
    text: "Amber",
  }),
  control: Object.freeze({
    alias: "Route control",
    tag: "ROUTE_CONTROL",
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

export const PPTX_TITLE_NOTES_FIXTURE = Object.freeze({
  presentationName: "launch-review.pptx",
  targetSlideName: "Go-no-go decision",
  untouchedSlideName: "Unchanged appendix",
  titleShapeName: "approval-title",
  originalTitle: "Decision: hold for legal review",
  replacementTitle: "Decision: approve controlled rollout",
  supportingText: "The scope, owner, and retained controls remain unchanged.",
  originalNotes: "Lead with the pending legal condition.\nClose with the accountable owner.",
  replacementNotes: "Lead with the approved controls.\nClose with the accountable rollout owner.",
  targetBackground: "#F1F5F9",
  untouchedBackground: "#FFF7ED",
});

// PromptBench keeps the ordinary title/notes example as a small plain-text
// workflow, while its ready presentation slice exercises the more important
// imported rich-notes boundary: a source-bound NotesSlide can change one
// existing ordinary run without flattening paragraphs, bullets, or sibling
// run formatting.
export const PPTX_RICH_NOTES_FIXTURE = Object.freeze({
  presentationName: "rich-notes-review.pptx",
  targetSlideName: "Go-no-go decision",
  untouchedSlideName: "Unchanged appendix",
  titleShapeName: "approval-title",
  originalTitle: "Decision: hold for legal review",
  replacementTitle: "Decision: approve controlled rollout",
  supportingText: "The speaker-note topology, visible layout, and appendix remain unchanged.",
  originalNotes: "Lead with the pending legal condition.\nClose with the accountable owner.",
  replacementNotes: "Lead with the approved control set.\nClose with the accountable owner.",
  originalNotesParagraphs: Object.freeze([
    Object.freeze({
      bulletCharacter: "•",
      runs: Object.freeze([
        Object.freeze({ text: "Lead with ", style: Object.freeze({ bold: true, fontSize: 18, fontFamily: "Aptos", color: "#0F172A" }) }),
        Object.freeze({ text: "the pending legal condition.", style: Object.freeze({ italic: true, fontSize: 18, color: "#7C2D12" }) }),
      ]),
    }),
    Object.freeze({
      autoNumber: Object.freeze({ type: "arabicPeriod", startAt: 2 }),
      runs: Object.freeze([
        Object.freeze({ text: "Close with the accountable owner.", style: Object.freeze({ fontSize: 16 }) }),
      ]),
    }),
  ]),
  targetRun: Object.freeze({
    paragraphIndex: 0,
    runIndex: 1,
    expectedText: "the pending legal condition.",
    replacementText: "the approved control set.",
    expectedStyle: Object.freeze({ italic: true, fontSize: 18, color: "#7c2d12" }),
    replacementStyle: Object.freeze({ bold: true, italic: false, fontSize: 18, color: "#0f766e" }),
  }),
  targetBackground: "#F1F5F9",
  untouchedBackground: "#FFF7ED",
});

// Reuse the same two-slide package for the non-visual source-bound rename
// case. The title, notes, direct backgrounds, and appendix are deliberate
// semantic and render canaries: the requested edit is only p:cSld/@name.
export const PPTX_SLIDE_NAME_FIXTURE = Object.freeze({
  presentationName: PPTX_TITLE_NOTES_FIXTURE.presentationName,
  expectedName: PPTX_TITLE_NOTES_FIXTURE.targetSlideName,
  replacementName: "Go decision: controlled rollout",
  untouchedSlideName: PPTX_TITLE_NOTES_FIXTURE.untouchedSlideName,
});

// This fixture gives PromptBench a semantic PowerPoint-section transaction
// that is deliberately invisible on the canvas. The four visible slides are
// package/render canaries while the only allowed package change is the
// canonical Office 2010 p14:sectionLst inside ppt/presentation.xml. Keeping
// the public facade IDs and native GUIDs fixed makes the full partition
// explicit: moving slide 2 from Context to Decision cannot be disguised as a
// partial one-section edit.
export const PPTX_SECTION_BOUNDARY_FIXTURE = Object.freeze({
  presentationName: "section-boundary-review.pptx",
  slides: Object.freeze([
    Object.freeze({ name: "Boundary opening", title: "1. Boundary opening", background: "#DBEAFE" }),
    Object.freeze({ name: "Boundary evidence", title: "2. Boundary evidence", background: "#DCFCE7" }),
    Object.freeze({ name: "Boundary decision", title: "3. Boundary decision", background: "#FEF3C7" }),
    Object.freeze({ name: "Boundary appendix", title: "4. Boundary appendix", background: "#FCE7F3" }),
  ]),
  sourceSections: Object.freeze([
    Object.freeze({
      id: "section/1",
      name: "Context",
      nativeId: "{01F07B81-39E6-4BBB-9B89-66EA253FBD29}",
      slideIds: Object.freeze(["presentation/slide/1", "presentation/slide/2"]),
    }),
    Object.freeze({
      id: "section/2",
      name: "Decision",
      nativeId: "{1FEF2C88-0CF2-4176-BA81-0DE6FD9D1274}",
      slideIds: Object.freeze(["presentation/slide/3"]),
    }),
    Object.freeze({
      id: "section/3",
      name: "Appendix",
      nativeId: "{2E92C0F3-07D0-4D22-8AC3-55C9651C42B1}",
      slideIds: Object.freeze(["presentation/slide/4"]),
    }),
  ]),
  replacementSections: Object.freeze([
    Object.freeze({
      id: "section/1",
      name: "Context",
      nativeId: "{01F07B81-39E6-4BBB-9B89-66EA253FBD29}",
      slideIds: Object.freeze(["presentation/slide/1"]),
    }),
    Object.freeze({
      id: "section/2",
      name: "Decision",
      nativeId: "{1FEF2C88-0CF2-4176-BA81-0DE6FD9D1274}",
      slideIds: Object.freeze(["presentation/slide/2", "presentation/slide/3"]),
    }),
    Object.freeze({
      id: "section/3",
      name: "Appendix",
      nativeId: "{2E92C0F3-07D0-4D22-8AC3-55C9651C42B1}",
      slideIds: Object.freeze(["presentation/slide/4"]),
    }),
  ]),
});

// This fixture exercises the narrow imported-slide clone profile rather than
// treating a presentation relationship graph as generally editable. Its
// source slide owns three accepted closed leaves: one canonical notes slide,
// one legacy comments XML leaf with a presentation-wide author catalog, and
// one literal-data chart whose ChartPart has no relationship graph. The
// appendix is a visible/package canary.
export const PPTX_CLOSED_LEAF_CLONE_FIXTURE = Object.freeze({
  presentationName: "release-review.pptx",
  sourceSlideName: "Release decision",
  appendixSlideName: "Appendix canary",
  sourceTitle: "Decision: approve controlled rollout",
  sourceSupportingText: "The original slide, notes, legacy comment, and appendix must remain unchanged.",
  sourceNotes: "Lead with the approved controls.\nClose with the accountable rollout owner.",
  sourceComment: "Confirm the original evidence before delivery.",
  commentAuthor: "Presentation Reviewer",
  commentCreated: "2026-07-18T03:05:00Z",
  chartTitle: "Control evidence by stage",
  chartCategories: Object.freeze(["Ready", "Watch", "Blocked"]),
  chartSeriesName: "Controls",
  chartValues: Object.freeze([68, 24, 8]),
  customShowName: "Board review route",
  customShowNativeId: 31,
  customShowText: "Open board review route",
  oleObjectName: "Embedded control evidence",
  oleWorkbookPart: "ppt/embeddings/release-control-evidence.xlsx",
  oleWorkbookRelationshipId: "rIdReleaseControlWorkbook",
  olePreviewPart: "ppt/media/release-control-evidence-preview.png",
  olePreviewRelationshipId: "rIdReleaseControlPreview",
  oleWorkbookMarker: "Release control evidence",
  sourceBackground: "#E0F2FE",
  appendixBackground: "#FEF3C7",
  appendixText: "Appendix: immutable evidence",
});

// This is one atomic safe-refusal boundary, not a claim that SmartArt, notes,
// and modern comments are all universally uneditable. The fourth slide has a
// canonical speaker-note and root/direct modern-comment thread as independent
// canaries. The first slide's SmartArt data part deliberately owns an external
// child relationship, which makes the diagram source-bound: a request that
// combines its node text with other edits must refuse as one transaction rather
// than partially mutating the notes/comment or flattening the diagram.
export const PPTX_SMARTART_NOTES_COMMENTS_BOUNDARY_FIXTURE = Object.freeze({
  presentationName: "strategy-review.pptx",
  slides: Object.freeze([
    Object.freeze({ name: "Strategy context", title: "1. Strategy context", background: "#DBEAFE" }),
    Object.freeze({ name: "Strategy evidence", title: "2. Strategy evidence", background: "#DCFCE7" }),
    Object.freeze({ name: "Strategy decision", title: "3. Strategy decision", background: "#FEF3C7" }),
    Object.freeze({ name: "Strategy review", title: "4. Strategy review", background: "#FCE7F3" }),
  ]),
  smartArt: Object.freeze({
    slideIndex: 0,
    name: "Strategy topology diagram",
    dataPartPath: "ppt/diagrams/strategy-data.xml",
    dataRelationshipPath: "ppt/diagrams/_rels/strategy-data.xml.rels",
    externalTarget: "https://example.invalid/strategy-review",
    nodes: Object.freeze([
      Object.freeze({ id: "{11111111-1111-4111-8111-111111111111}", text: "Observe" }),
      Object.freeze({ id: "{22222222-2222-4222-8222-222222222222}", text: "Decide" }),
      Object.freeze({ id: "{33333333-3333-4333-8333-333333333333}", text: "Scale candidate" }),
    ]),
  }),
  notes: Object.freeze({
    slideIndex: 3,
    text: "Fourth-slide speaker-notes canary: preserve the review context.",
  }),
  comment: Object.freeze({
    slideIndex: 3,
    targetName: "Strategy comment target",
    targetText: "Review this controlled strategy.",
    root: Object.freeze({
      id: "{44444444-4444-4444-8444-444444444444}",
      personId: "{AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}",
      author: "Review Owner",
      initials: "RO",
      userId: "review.owner@example.test",
      created: "2026-07-28T08:00:00Z",
      text: "Confirm the controlled strategy boundary.",
    }),
    directReply: Object.freeze({
      id: "{55555555-5555-4555-8555-555555555555}",
      personId: "{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}",
      author: "Evidence Owner",
      initials: "EO",
      userId: "evidence.owner@example.test",
      created: "2026-07-28T08:05:00Z",
      text: "Evidence is attached for the review.",
    }),
  }),
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

function appendBeforeFinalElement(xml, localName, fragment, label) {
  const closing = new RegExp(`</(?:[A-Za-z_][\\w.-]*:)?${localName}\\s*>`, "ig");
  const matches = [...String(xml).matchAll(closing)];
  const match = matches.at(-1);
  if (!match || match.index == null) throw new Error(`XLSX threaded-comment fixture is missing ${label}.`);
  return `${xml.slice(0, match.index)}${fragment}${xml.slice(match.index)}`;
}

function rootElementPrefix(xml, localName, label) {
  const match = new RegExp(`<([A-Za-z_][\\w.-]*:)?${localName}\\b`, "i").exec(String(xml));
  if (!match) throw new Error(`XLSX threaded-comment fixture is missing ${label}.`);
  return match[1] || "";
}

async function appendNestedXlsxThreadedReply(bytes, fixture) {
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
  const threadedPaths = paths.filter((name) => /^xl\/threadedcomments\/[^/]+\.xml$/i.test(name));
  const personPaths = paths.filter((name) => /^xl\/persons\/[^/]+\.xml$/i.test(name));
  if (threadedPaths.length !== 1 || personPaths.length !== 1) {
    throw new Error("XLSX nested threaded-comment fixture requires one threaded-comments and one person part.");
  }
  const [threadedPath] = threadedPaths;
  const [personPath] = personPaths;
  const [threadedXml, personXml] = await Promise.all([
    zip.file(threadedPath)?.async("text"),
    zip.file(personPath)?.async("text"),
  ]);
  if (!threadedXml || !personXml) throw new Error("XLSX nested threaded-comment fixture could not read native comment parts.");
  const nested = fixture.nestedReply;
  const threadedPrefix = rootElementPrefix(threadedXml, "ThreadedComments", "ThreadedComments root");
  const personPrefix = rootElementPrefix(personXml, "personList", "personList root");
  const comment = `<${threadedPrefix}threadedComment ref="${fixture.address}" dT="${nested.date}" personId="${nested.personId}" id="${nested.id}" parentId="${fixture.directReply.id}" done="0"><${threadedPrefix}text>${xmlEscape(nested.text)}</${threadedPrefix}text></${threadedPrefix}threadedComment>`;
  const person = `<${personPrefix}person displayName="${xmlEscape(nested.author)}" id="${nested.personId}" userId="${xmlEscape(nested.userId)}" providerId="${xmlEscape(nested.providerId)}"/>`;
  zip.file(threadedPath, appendBeforeFinalElement(threadedXml, "ThreadedComments", comment, "ThreadedComments root"));
  zip.file(personPath, appendBeforeFinalElement(personXml, "personList", person, "personList root"));
  return new Uint8Array(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } }));
}

export async function generateXlsxThreadedNestedReplyBoundary(target) {
  const fixture = XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE;
  await generateXlsxThreadedReview(target);
  const nestedBytes = await appendNestedXlsxThreadedReply(await fs.readFile(target), fixture);
  await fs.writeFile(target, nestedBytes);
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

function appendBeforeClosingTag(xml, closingTag, fragment, label) {
  const index = xml.lastIndexOf(closingTag);
  if (index < 0) throw new Error(`DOCX modern-comment fixture is missing ${label}.`);
  return `${xml.slice(0, index)}${fragment}${xml.slice(index)}`;
}

async function appendNestedModernCommentReply(authored, fixture) {
  const [, directReply, nestedReply] = fixture.comments;
  const zip = await JSZip.loadAsync(await authored.arrayBuffer());
  const requiredParts = [
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsIds.xml",
    "word/commentsExtensible.xml",
    "word/people.xml",
  ];
  const contents = Object.fromEntries(await Promise.all(requiredParts.map(async (partPath) => {
    const part = zip.file(partPath);
    if (!part) throw new Error(`DOCX modern-comment fixture did not author ${partPath}.`);
    return [partPath, await part.async("text")];
  })));
  const commentBody = `<w:comment w:id="${nestedReply.nativeId}" w:author="${xmlEscape(nestedReply.author)}" w:initials="${xmlEscape(nestedReply.initials)}" w:date="${nestedReply.date}"><w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="${nestedReply.paraId}"><w:r><w:t>${xmlEscape(nestedReply.text)}</w:t></w:r></w:p></w:comment>`;
  contents["word/comments.xml"] = appendBeforeClosingTag(contents["word/comments.xml"], "</w:comments>", commentBody, "w:comments");
  contents["word/commentsExtended.xml"] = appendBeforeClosingTag(
    contents["word/commentsExtended.xml"],
    "</w15:commentsEx>",
    `<w15:commentEx w15:paraId="${nestedReply.paraId}" w15:paraIdParent="${directReply.paraId}" w15:done="0"/>`,
    "w15:commentsEx",
  );
  contents["word/commentsIds.xml"] = appendBeforeClosingTag(
    contents["word/commentsIds.xml"],
    "</w16cid:commentsIds>",
    `<w16cid:commentId w16cid:paraId="${nestedReply.paraId}" w16cid:durableId="${nestedReply.durableId}"/>`,
    "w16cid:commentsIds",
  );
  contents["word/commentsExtensible.xml"] = appendBeforeClosingTag(
    contents["word/commentsExtensible.xml"],
    "</w16cex:commentsExtensible>",
    `<w16cex:commentExtensible w16cex:durableId="${nestedReply.durableId}" w16cex:dateUtc="${nestedReply.dateUtc}"/>`,
    "w16cex:commentsExtensible",
  );
  contents["word/people.xml"] = appendBeforeClosingTag(
    contents["word/people.xml"],
    "</w15:people>",
    `<w15:person w15:author="${xmlEscape(nestedReply.author)}"><w15:presenceInfo w15:providerId="${xmlEscape(nestedReply.providerId)}" w15:userId="${xmlEscape(nestedReply.userId)}"/></w15:person>`,
    "w15:people",
  );
  for (const [partPath, xml] of Object.entries(contents)) zip.file(partPath, xml);
  return new Uint8Array(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } }));
}

export async function generateDocxModernCommentReplyBoundary(target) {
  const fixture = DOCX_MODERN_COMMENT_REPLY_BOUNDARY_FIXTURE;
  const [root, directReply] = fixture.comments;
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
  document.addParagraph(fixture.supportingText, { paragraphFormat: { spaceAfterTwips: 120 } });
  const rootComment = document.addComment(decision, root.text, {
    author: root.author,
    initials: root.initials,
    date: root.date,
    resolved: root.resolved,
    paraId: root.paraId,
    durableId: root.durableId,
    dateUtc: root.dateUtc,
    person: { providerId: root.providerId, userId: root.userId },
  });
  document.replyToComment(rootComment, directReply.text, {
    author: directReply.author,
    initials: directReply.initials,
    date: directReply.date,
    paraId: directReply.paraId,
    durableId: directReply.durableId,
    dateUtc: directReply.dateUtc,
    person: { providerId: directReply.providerId, userId: directReply.userId },
  });
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated DOCX modern-comment boundary fixture failed model verification: " + verification.ndjson);
  const authored = await DocumentFile.exportDocx(document);
  const nestedBytes = await appendNestedModernCommentReply(authored, fixture);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, nestedBytes);
  return { path: target, type: DOCX_MIME };
}

function wordParagraph(text) {
  return `<w:p><w:r><w:t>${xmlEscape(text)}</w:t></w:r></w:p>`;
}

function wordTableCell(content, properties = "") {
  return `<w:tc><w:tcPr><w:tcW w:w="2800" w:type="dxa"/>${properties}</w:tcPr>${content}</w:tc>`;
}

function wordTableRow(cells) {
  return `<w:tr>${cells.join("")}</w:tr>`;
}

function complexTableTopologyXml(fixture) {
  const nestedTable = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="1300"/><w:gridCol w:w="2700"/></w:tblGrid>${wordTableRow([
    wordTableCell(wordParagraph(fixture.nestedSchedule[0])),
    wordTableCell(wordParagraph(fixture.nestedSchedule[1])),
  ])}</w:tbl>`;
  const routeControl = `<w:sdt><w:sdtPr><w:alias w:val="${xmlEscape(fixture.control.alias)}"/><w:tag w:val="${xmlEscape(fixture.control.tag)}"/><w:text/></w:sdtPr><w:sdtContent>${wordParagraph(fixture.routeValue)}</w:sdtContent></w:sdt><w:p/>`;
  const revisedState = `<w:p><w:ins w:id="${fixture.revision.id}" w:author="${xmlEscape(fixture.revision.author)}" w:date="${fixture.revision.date}"><w:r><w:t>${xmlEscape(fixture.revision.text)}</w:t></w:r></w:ins></w:p>`;
  return `<w:tbl><w:tblPr><w:tblStyle w:val="${fixture.styleId}"/><w:tblW w:w="8400" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCaption w:val="${fixture.tableCaption}"/></w:tblPr><w:tblGrid><w:gridCol w:w="2200"/><w:gridCol w:w="3100"/><w:gridCol w:w="3100"/></w:tblGrid>${wordTableRow(fixture.headers.map((header) => wordTableCell(wordParagraph(header))))}${wordTableRow([
    wordTableCell(wordParagraph(fixture.mergedDose), '<w:vMerge w:val="restart"/>'),
    wordTableCell(routeControl),
    wordTableCell(revisedState),
  ])}${wordTableRow([
    wordTableCell("<w:p/>", "<w:vMerge/>"),
    wordTableCell(`${wordParagraph(fixture.nestedScheduleLabel)}${nestedTable}<w:p/>`),
    wordTableCell(wordParagraph("Pending")),
  ])}</w:tbl>`;
}

function insertBeforeFinalDocumentSectionProperties(documentXml, fragment) {
  const index = documentXml.lastIndexOf("<w:sectPr");
  if (index < 0) throw new Error("DOCX complex-table fixture is missing the final w:sectPr.");
  return `${documentXml.slice(0, index)}${fragment}${documentXml.slice(index)}`;
}

function customClinicalTableStyleXml(fixture) {
  return `<w:style w:type="table" w:styleId="${fixture.styleId}" w:customStyle="1"><w:name w:val="Clinical protocol table"/><w:uiPriority w:val="99"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="8" w:space="0" w:color="2F5597"/><w:left w:val="single" w:sz="8" w:space="0" w:color="2F5597"/><w:bottom w:val="single" w:sz="8" w:space="0" w:color="2F5597"/><w:right w:val="single" w:sz="8" w:space="0" w:color="2F5597"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/></w:tblBorders></w:tblPr></w:style>`;
}

async function appendComplexTableTopology(authored, fixture) {
  const zip = await JSZip.loadAsync(await authored.arrayBuffer());
  const requiredParts = ["word/document.xml", "word/styles.xml"];
  const contents = Object.fromEntries(await Promise.all(requiredParts.map(async (partPath) => {
    const part = zip.file(partPath);
    if (!part) throw new Error(`DOCX complex-table fixture did not author ${partPath}.`);
    return [partPath, await part.async("text")];
  })));
  contents["word/document.xml"] = insertBeforeFinalDocumentSectionProperties(contents["word/document.xml"], complexTableTopologyXml(fixture));
  contents["word/styles.xml"] = appendBeforeClosingTag(contents["word/styles.xml"], "</w:styles>", customClinicalTableStyleXml(fixture), "w:styles");
  for (const [partPath, xml] of Object.entries(contents)) zip.file(partPath, xml);
  return new Uint8Array(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } }));
}

export async function generateDocxComplexTableTopologyBoundary(target) {
  const fixture = DOCX_COMPLEX_TABLE_TOPOLOGY_BOUNDARY_FIXTURE;
  const document = DocumentModel.create({
    name: fixture.title,
    defaultRunStyle: { fontFamily: "Aptos", fontSize: 11, color: "#172033" },
    blocks: [],
  });
  document.applyDesignPreset("report");
  document.addParagraph(fixture.title, {
    paragraphFormat: { spaceAfterTwips: 160 },
    runs: [{ text: fixture.title, style: { bold: true, fontSize: 16, color: "#123B5D" } }],
  });
  document.addTable({
    name: "baseline-clinical-review",
    values: fixture.baselineTable,
    widthDxa: 6000,
    columnWidthsDxa: [3600, 2400],
    headerFill: "DCE6F1",
  });
  document.addParagraph(fixture.introduction, { paragraphFormat: { spaceBeforeTwips: 120, spaceAfterTwips: 120 } });
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated DOCX complex-table boundary fixture failed model verification: " + verification.ndjson);
  const authored = await DocumentFile.exportDocx(document);
  const boundaryBytes = await appendComplexTableTopology(authored, fixture);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, boundaryBytes);
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

export async function generatePptxTitleNotesReview(target) {
  const fixture = PPTX_TITLE_NOTES_FIXTURE;
  const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  const decision = presentation.slides.add({ name: fixture.targetSlideName });
  decision.setBackground({ fill: fixture.targetBackground, mode: "solid" });
  const title = decision.shapes.add({
    name: fixture.titleShapeName,
    geometry: "textbox",
    position: { left: 72, top: 72, width: 1040, height: 96 },
    text: fixture.originalTitle,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  title.text.style = { fontSize: 34, bold: true, color: "#0F172A" };
  const supporting = decision.shapes.add({
    name: "supporting-copy",
    geometry: "textbox",
    position: { left: 72, top: 194, width: 880, height: 80 },
    text: fixture.supportingText,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  supporting.text.style = { fontSize: 18, color: "#334155" };
  decision.addNotes(fixture.originalNotes);

  const appendix = presentation.slides.add({ name: fixture.untouchedSlideName });
  appendix.setBackground({ fill: fixture.untouchedBackground, mode: "solid" });
  const appendixTitle = appendix.shapes.add({
    name: "appendix-title",
    geometry: "textbox",
    position: { left: 72, top: 72, width: 900, height: 96 },
    text: "Appendix: unchanged evidence",
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  appendixTitle.text.style = { fontSize: 30, bold: true, color: "#7C2D12" };

  const verification = presentation.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated PPTX title/notes fixture failed model verification: " + verification.ndjson);
  const exported = await PresentationFile.exportPptx(presentation);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await exported.arrayBuffer()));
  return { path: target, type: PPTX_MIME };
}

export async function generatePptxRichNotesReview(target) {
  const fixture = PPTX_RICH_NOTES_FIXTURE;
  const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  const decision = presentation.slides.add({ name: fixture.targetSlideName });
  decision.setBackground({ fill: fixture.targetBackground, mode: "solid" });
  const title = decision.shapes.add({
    name: fixture.titleShapeName,
    geometry: "textbox",
    position: { left: 72, top: 72, width: 1040, height: 96 },
    text: fixture.originalTitle,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  title.text.style = { fontSize: 34, bold: true, color: "#0F172A" };
  const supporting = decision.shapes.add({
    name: "supporting-copy",
    geometry: "textbox",
    position: { left: 72, top: 194, width: 880, height: 80 },
    text: fixture.supportingText,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  supporting.text.style = { fontSize: 18, color: "#334155" };
  decision.addNotes(fixture.originalNotesParagraphs);

  const appendix = presentation.slides.add({ name: fixture.untouchedSlideName });
  appendix.setBackground({ fill: fixture.untouchedBackground, mode: "solid" });
  const appendixTitle = appendix.shapes.add({
    name: "appendix-title",
    geometry: "textbox",
    position: { left: 72, top: 72, width: 900, height: 96 },
    text: "Appendix: unchanged evidence",
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  appendixTitle.text.style = { fontSize: 30, bold: true, color: "#7C2D12" };

  const verification = presentation.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated PPTX rich-notes fixture failed model verification: " + verification.ndjson);
  const exported = await PresentationFile.exportPptx(presentation);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await exported.arrayBuffer()));
  return { path: target, type: PPTX_MIME };
}

export async function generatePptxSlideNameReview(target) {
  return generatePptxTitleNotesReview(target);
}

export async function generatePptxSectionBoundaryReview(target) {
  const fixture = PPTX_SECTION_BOUNDARY_FIXTURE;
  const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  const slides = fixture.slides.map((definition, index) => {
    const slide = presentation.slides.add({ name: definition.name });
    slide.setBackground({ fill: definition.background, mode: "solid" });
    const title = slide.shapes.add({
      name: `section-boundary-title-${index + 1}`,
      geometry: "textbox",
      position: { left: 72, top: 72, width: 1000, height: 96 },
      text: definition.title,
      fill: "none",
      line: { style: "solid", fill: "none", width: 0 },
    });
    title.text.style = { fontSize: 34, bold: true, color: "#0F172A" };
    const supporting = slide.shapes.add({
      name: `section-boundary-canary-${index + 1}`,
      geometry: "textbox",
      position: { left: 72, top: 194, width: 920, height: 80 },
      text: "Visible geometry and all non-section package parts must remain unchanged.",
      fill: "none",
      line: { style: "solid", fill: "none", width: 0 },
    });
    supporting.text.style = { fontSize: 18, color: "#334155" };
    return slide;
  });
  for (const section of fixture.sourceSections) {
    presentation.sections.add({
      name: section.name,
      nativeId: section.nativeId,
      slides: section.slideIds.map((id) => slides[Number(id.split("/").at(-1)) - 1]),
    });
  }
  // Source-free model IDs are allocator-local (`sl/...`). The evaluator's
  // public fixture IDs intentionally become stable `presentation/slide/N`
  // only after the PPTX is imported, so translate the expected membership for
  // this pre-export self-check rather than asserting an accidental allocator
  // identity.
  const expectedModelSections = fixture.sourceSections.map((section) => ({
    ...section,
    slideIds: section.slideIds.map((id) => slides[Number(id.split("/").at(-1)) - 1].id),
  }));
  const actualSections = presentation.sections.items.map((section) => section.toJSON());
  if (JSON.stringify(actualSections) !== JSON.stringify(expectedModelSections)) {
    throw new Error("Generated PPTX section-boundary fixture did not retain the fixed source partition.");
  }
  const verification = presentation.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated PPTX section-boundary fixture failed model verification: " + verification.ndjson);
  const exported = await PresentationFile.exportPptx(presentation);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await exported.arrayBuffer()));
  return { path: target, type: PPTX_MIME };
}

async function addClosedCloneOleWorkbook(exported, fixture) {
  const embeddedWorkbook = Workbook.create();
  embeddedWorkbook.worksheets.add("Evidence").getRange("A1:B3").values = [
    [fixture.oleWorkbookMarker, null],
    ["Control", "Status"],
    ["Release gate", "Approved"],
  ];
  const embeddedWorkbookFile = await SpreadsheetFile.exportXlsx(embeddedWorkbook);
  const zip = await JSZip.loadAsync(exported.bytes);
  const [slideXml, slideRelationships] = await Promise.all([
    zip.file("ppt/slides/slide1.xml").async("text"),
    zip.file("ppt/slides/_rels/slide1.xml.rels").async("text"),
  ]);
  const oleFrame = `<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="100" name="${fixture.oleObjectName}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="10191750" y="3238500"/><a:ext cx="1524000" cy="1143000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj showAsIcon="1" r:id="${fixture.oleWorkbookRelationshipId}" imgW="965200" imgH="609600" progId="Excel.Sheet.12"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${fixture.olePreviewRelationshipId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="10191750" y="3238500"/><a:ext cx="1524000" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>`;
  return PresentationFile.patchPptx(exported, [
    { path: "ppt/slides/slide1.xml", xml: slideXml.replace("</p:spTree>", `${oleFrame}</p:spTree>`) },
    { path: "ppt/slides/_rels/slide1.xml.rels", xml: slideRelationships.replace("</Relationships>", `<Relationship Id="${fixture.oleWorkbookRelationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/release-control-evidence.xlsx"/><Relationship Id="${fixture.olePreviewRelationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/release-control-evidence-preview.png"/></Relationships>`) },
    { path: fixture.oleWorkbookPart, bytes: embeddedWorkbookFile.bytes, contentType: XLSX_MIME },
    { path: fixture.olePreviewPart, bytes: Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", "base64"), contentType: "image/png" },
  ]);
}

export async function generatePptxClosedLeafClone(target) {
  const fixture = PPTX_CLOSED_LEAF_CLONE_FIXTURE;
  const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  const source = presentation.slides.add({ name: fixture.sourceSlideName });
  source.setBackground({ fill: fixture.sourceBackground, mode: "solid" });
  const title = source.shapes.add({
    name: "release-title",
    geometry: "textbox",
    position: { left: 72, top: 72, width: 1040, height: 96 },
    text: fixture.sourceTitle,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  title.text.style = { fontSize: 34, bold: true, color: "#0C4A6E" };
  const supporting = source.shapes.add({
    name: "release-supporting-copy",
    geometry: "textbox",
    position: { left: 72, top: 194, width: 920, height: 88 },
    text: fixture.sourceSupportingText,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  supporting.text.style = { fontSize: 18, color: "#334155" };
  const route = source.shapes.add({
    name: "board-route-link",
    geometry: "textbox",
    position: { left: 1048, top: 194, width: 184, height: 108 },
    text: [{ runs: [{
      text: fixture.customShowText,
      link: { customShow: fixture.customShowName, returnToSlide: true, tooltip: "Open the board route" },
    }] }],
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  route.text.style = { fontSize: 15, bold: true, color: "#0369A1" };
  source.addNotes(fixture.sourceNotes);
  source.comments.addThread(undefined, fixture.sourceComment, {
    author: fixture.commentAuthor,
    created: fixture.commentCreated,
    position: { x: 360, y: 240 },
  });
  source.charts.add("bar", {
    name: "release-evidence-chart",
    position: { left: 72, top: 318, width: 980, height: 320 },
    title: fixture.chartTitle,
    categories: [...fixture.chartCategories],
    series: [{ name: fixture.chartSeriesName, values: [...fixture.chartValues], fill: "#0284C7" }],
    axes: {
      category: { title: "Evidence stage" },
      value: { title: "Share", min: 0, max: 80, majorUnit: 20 },
    },
    legend: false,
    dataLabels: { showValue: true, position: "top" },
  });

  const appendix = presentation.slides.add({ name: fixture.appendixSlideName });
  appendix.setBackground({ fill: fixture.appendixBackground, mode: "solid" });
  const appendixTitle = appendix.shapes.add({
    name: "appendix-title",
    geometry: "textbox",
    position: { left: 72, top: 72, width: 900, height: 96 },
    text: fixture.appendixText,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  appendixTitle.text.style = { fontSize: 30, bold: true, color: "#92400E" };
  presentation.customShows.add({
    name: fixture.customShowName,
    nativeId: fixture.customShowNativeId,
    slides: [source, appendix],
  });

  const verification = presentation.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated PPTX closed-leaf clone fixture failed model verification: " + verification.ndjson);
  const exported = await PresentationFile.exportPptx(presentation);
  const patchedSource = await addClosedCloneOleWorkbook(exported, fixture);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await patchedSource.arrayBuffer()));
  return { path: target, type: PPTX_MIME };
}

async function addConnectedSmartArtBoundary(exported, fixture) {
  const smartArt = fixture.smartArt;
  const slideNumber = smartArt.slideIndex + 1;
  const slidePath = `ppt/slides/slide${slideNumber}.xml`;
  const relationshipsPath = `ppt/slides/_rels/slide${slideNumber}.xml.rels`;
  const zip = await JSZip.loadAsync(exported.bytes);
  const [slideXml, relationshipsXml] = await Promise.all([
    zip.file(slidePath)?.async("text"),
    zip.file(relationshipsPath)?.async("text"),
  ]);
  if (!slideXml?.includes("</p:spTree>") || !relationshipsXml?.includes("</Relationships>")) {
    throw new Error("PPTX SmartArt boundary fixture could not locate the source slide tree or relationships root.");
  }
  const frame = `<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="120" name="${xmlEscape(smartArt.name)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="1828800"/><a:ext cx="5486400" cy="2743200"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdStrategyData" r:lo="rIdStrategyLayout" r:qs="rIdStrategyStyle" r:cs="rIdStrategyColors"/></a:graphicData></a:graphic></p:graphicFrame>`;
  const relationships = '<Relationship Id="rIdStrategyData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/strategy-data.xml"/><Relationship Id="rIdStrategyLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/strategy-layout.xml"/><Relationship Id="rIdStrategyStyle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/strategy-style.xml"/><Relationship Id="rIdStrategyColors" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/strategy-colors.xml"/>';
  const nodes = smartArt.nodes.map((node) => `<dgm:pt modelId="${node.id}" type="doc"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xmlEscape(node.text)}</a:t></a:r></a:p></dgm:t></dgm:pt>`).join("");
  return PresentationFile.patchPptx(exported, [
    { path: slidePath, xml: slideXml.replace("</p:spTree>", `${frame}</p:spTree>`) },
    { path: relationshipsPath, xml: relationshipsXml.replace("</Relationships>", `${relationships}</Relationships>`) },
    {
      path: smartArt.dataPartPath,
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
      xml: `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst>${nodes}</dgm:ptLst><dgm:cxnLst/><dgm:bg/><dgm:whole/></dgm:dataModel>`,
    },
    {
      path: "ppt/diagrams/strategy-layout.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
      xml: '<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:office-kit:strategy-layout"><dgm:title val="Strategy"/><dgm:desc val="Strategy layout"/><dgm:catLst/><dgm:layoutNode name="root"/></dgm:layoutDef>',
    },
    {
      path: "ppt/diagrams/strategy-style.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
      xml: '<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:office-kit:strategy-style"><dgm:title val="Strategy"/><dgm:desc val="Strategy style"/><dgm:catLst/><dgm:styleLbl name="node0"/></dgm:styleDef>',
    },
    {
      path: "ppt/diagrams/strategy-colors.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
      xml: '<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:office-kit:strategy-colors"><dgm:title val="Strategy"/><dgm:desc val="Strategy colors"/><dgm:catLst/></dgm:colorsDef>',
    },
    {
      path: smartArt.dataRelationshipPath,
      xml: `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdExternalStrategyLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xmlEscape(smartArt.externalTarget)}" TargetMode="External"/></Relationships>`,
    },
  ]);
}

export async function generatePptxSmartArtNotesCommentsBoundary(target) {
  const fixture = PPTX_SMARTART_NOTES_COMMENTS_BOUNDARY_FIXTURE;
  const presentation = Presentation.create({
    slideSize: { width: 1280, height: 720 },
    commentFormat: "modern",
  });
  const slides = fixture.slides.map((definition, index) => {
    const slide = presentation.slides.add({ name: definition.name });
    slide.setBackground({ fill: definition.background, mode: "solid" });
    const title = slide.shapes.add({
      id: `strategy-boundary-title-${index + 1}`,
      name: `strategy-boundary-title-${index + 1}`,
      geometry: "textbox",
      position: { left: 72, top: 72, width: 1000, height: 96 },
      text: definition.title,
      fill: "none",
      line: { style: "solid", fill: "none", width: 0 },
    });
    title.text.style = { fontSize: 34, bold: true, color: "#0F172A" };
    const supporting = slide.shapes.add({
      name: `strategy-boundary-canary-${index + 1}`,
      geometry: "textbox",
      // The opaque SmartArt graphic frame is injected at 192–480px on slide
      // one. Keep this ordinary text canary below it so a public import and
      // visual verification of the locked source is genuinely clean.
      position: { left: 72, top: index === fixture.smartArt.slideIndex ? 520 : 194, width: 920, height: 80 },
      text: "Preserve this source-bound review package without partial edits.",
      fill: "none",
      line: { style: "solid", fill: "none", width: 0 },
    });
    supporting.text.style = { fontSize: 18, color: "#334155" };
    return slide;
  });
  const noteSlide = slides[fixture.notes.slideIndex];
  noteSlide.addNotes(fixture.notes.text);
  const targetShape = noteSlide.shapes.add({
    id: "strategy-comment-target",
    name: fixture.comment.targetName,
    geometry: "textbox",
    position: { left: 72, top: 314, width: 920, height: 72 },
    text: fixture.comment.targetText,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  targetShape.text.style = { fontSize: 18, color: "#334155" };
  const root = fixture.comment.root;
  const thread = noteSlide.comments.addThread({
    textMatch: { element: targetShape, query: fixture.comment.targetText, occurrence: 0 },
  }, root.text, {
    id: root.id,
    author: root.author,
    created: root.created,
    nativeFormat: "modern",
    position: { x: 1_234_500, y: 2_345_600, unit: "emu" },
    comments: [{
      nativeId: root.id,
      author: root.author,
      person: {
        id: root.personId,
        name: root.author,
        initials: root.initials,
        userId: root.userId,
        providerId: "None",
      },
      text: root.text,
      created: root.created,
      status: "active",
    }],
  });
  const reply = fixture.comment.directReply;
  thread.addReply(reply.text, {
    nativeId: reply.id,
    author: reply.author,
    person: {
      id: reply.personId,
      name: reply.author,
      initials: reply.initials,
      userId: reply.userId,
      providerId: "None",
    },
    created: reply.created,
    status: "active",
  });
  const verification = presentation.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated PPTX SmartArt boundary fixture failed model verification: " + verification.ndjson);
  const exported = await PresentationFile.exportPptx(presentation);
  const patched = await addConnectedSmartArtBoundary(exported, fixture);
  const imported = await PresentationFile.importPptx(patched);
  const importedVerification = imported.verify({ visualQa: true });
  const diagram = imported.slides.getItem(fixture.smartArt.slideIndex).nativeObjects.items
    .find((item) => item.name === fixture.smartArt.name);
  const reviewSlide = imported.slides.getItem(fixture.notes.slideIndex);
  const comment = reviewSlide.comments.items[0];
  if (!diagram || diagram.nativeKind !== "diagram" || diagram.diagramText !== undefined
    || diagram.inspectRecord().nativeParts.find((part) => part.path === fixture.smartArt.dataPartPath)?.relationships !== 1
    || reviewSlide.speakerNotes.text !== fixture.notes.text
    || reviewSlide.comments.items.length !== 1
    || comment?.comments.length !== 2
    || comment?.comments[0]?.text !== root.text
    || comment?.comments[1]?.text !== reply.text
    || !importedVerification.ok) {
    throw new Error("Generated PPTX SmartArt boundary fixture did not reimport its source-bound diagram and review canaries.");
  }
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await patched.arrayBuffer()));
  return { path: target, type: PPTX_MIME };
}

export async function generateOfficeInput(generator, target) {
  if (generator === "xlsx-threaded-review") return generateXlsxThreadedReview(target);
  if (generator === "xlsx-threaded-nested-reply-boundary") return generateXlsxThreadedNestedReplyBoundary(target);
  if (generator === "xlsx-growth-update") return generateXlsxGrowthUpdate(target);
  if (generator === "xlsx-connection-refresh") return generateXlsxConnectionRefresh(target);
  if (generator === "xlsx-pivot-refresh") return generateXlsxPivotRefresh(target);
  if (generator === "docx-classic-comment-review") return generateDocxClassicCommentReview(target);
  if (generator === "docx-modern-comment-reply-boundary") return generateDocxModernCommentReplyBoundary(target);
  if (generator === "docx-complex-table-topology-boundary") return generateDocxComplexTableTopologyBoundary(target);
  if (generator === "docx-header-text-review") return generateDocxHeaderTextReview(target);
  if (generator === "docx-footer-text-review") return generateDocxFooterTextReview(target);
  if (generator === "docx-section-page-numbering-review") return generateDocxSectionPageNumberingReview(target);
  if (generator === "pptx-title-notes-review") return generatePptxTitleNotesReview(target);
  if (generator === "pptx-rich-notes-review") return generatePptxRichNotesReview(target);
  if (generator === "pptx-slide-name-review") return generatePptxSlideNameReview(target);
  if (generator === "pptx-section-boundary-review") return generatePptxSectionBoundaryReview(target);
  if (generator === "pptx-closed-leaf-clone") return generatePptxClosedLeafClone(target);
  if (generator === "pptx-smartart-notes-comments-boundary") return generatePptxSmartArtNotesCommentsBoundary(target);
  return null;
}
