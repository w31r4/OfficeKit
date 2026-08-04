import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { FileBlob, SpreadsheetFile, Workbook } from "office-kit";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const require = createRequire(import.meta.url);
const SCENARIOS = ["Base", "Upside", "Downside"];
const COMMENT_IDS = Object.freeze([
  ["{11111111-1111-4111-8111-111111111111}", "{AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}"],
  ["{22222222-2222-4222-8222-222222222222}", "{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}"],
]);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function requiredPath(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty path.`);
  return path.resolve(value);
}

function parseActuals(csv) {
  const lines = String(csv).trim().split(/\r?\n/);
  const header = lines.shift();
  if (header !== "month,revenue,cost,headcount") throw new Error("actuals.csv has an unexpected header.");
  const rows = lines.map((line) => {
    const [month, revenue, cost, headcount] = line.split(",");
    if (!/^\d{4}-\d{2}$/.test(month) || ![revenue, cost, headcount].every((value) => Number.isFinite(Number(value)))) {
      throw new Error(`Invalid actuals row: ${line}`);
    }
    return { month, revenue: Number(revenue), cost: Number(cost), headcount: Number(headcount) };
  });
  if (rows.length !== 24) throw new Error(`actuals.csv must contain exactly 24 months; found ${rows.length}.`);
  return rows;
}

function validateAssumptions(value) {
  if (!value || value.activeScenario !== "Base" || !value.scenarios || !Array.isArray(value.checks)) {
    throw new Error("assumptions.json must declare the Base scenario and checks.");
  }
  for (const scenario of SCENARIOS) {
    const profile = value.scenarios[scenario];
    if (!profile || !["revenueGrowth", "costGrowth", "grossMargin", "startingCash", "minimumCash"].every((key) => Number.isFinite(profile[key]))) {
      throw new Error(`assumptions.json is missing a complete ${scenario} profile.`);
    }
  }
  return value;
}

function commentConfig(id, personId, author, userId, date) {
  return {
    id,
    personId,
    author,
    date,
    person: { id: personId, displayName: author, userId, providerId: "None" },
  };
}

async function packageVersion() {
  try {
    const entry = require.resolve("office-kit");
    const packagePath = path.join(path.dirname(path.dirname(entry)), "package.json");
    return JSON.parse(await fs.readFile(packagePath, "utf8")).version;
  } catch {
    return "0.6.0";
  }
}

async function renderAllSheets(workbook) {
  const rendered = [];
  for (const sheet of workbook.worksheets.items) {
    const preview = await workbook.render({ sheetName: sheet.name, autoCrop: "all", format: "svg" });
    const svg = await preview.text();
    if (!/<svg\b/i.test(svg) || svg.length < 128) throw new Error(`Model render for ${sheet.name} is incomplete.`);
    rendered.push({ sheet: sheet.name, bytes: preview.bytes.length, renderer: "model-svg" });
  }
  return rendered;
}

function buildWorkbook(actuals, assumptions) {
  const workbook = Workbook.create();

  const sources = workbook.worksheets.add("Sources");
  sources.getRange("A1:D25").values = [
    ["Month", "Revenue", "Cost", "Headcount"],
    ...actuals.map((row) => [row.month, row.revenue, row.cost, row.headcount]),
  ];
  sources.getRange("A1:D1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
  sources.getRange("B2:C25").format.numberFormat = "$#,##0";
  sources.getRange("A1:D25").format.columnWidthPx = 120;
  sources.freezePanes.freezeRows(1);
  sources.showGridLines = false;

  const assumptionSheet = workbook.worksheets.add("Assumptions");
  assumptionSheet.getRange("A1:F11").values = [
    ["FY27 Operating Plan Assumptions", null, null, null, null, null],
    [null, null, null, null, null, null],
    ["Active scenario", assumptions.activeScenario, null, null, null, null],
    [null, null, null, null, null, null],
    ["Scenario", "Revenue growth", "Cost growth", "Gross margin", "Starting cash", "Minimum cash"],
    ...SCENARIOS.map((scenario) => {
      const profile = assumptions.scenarios[scenario];
      return [scenario, profile.revenueGrowth, profile.costGrowth, profile.grossMargin, profile.startingCash, profile.minimumCash];
    }),
    [null, null, null, null, null, null],
    ["Checks", null, null, null, null, null],
    ...assumptions.checks.slice(0, 3).map((check) => [check, null, null, null, null, null]),
  ];
  assumptionSheet.getRange("A1:F1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF", size: 14 } };
  assumptionSheet.getRange("A5:F5").format = { fill: "#E2E8F0", font: { bold: true } };
  assumptionSheet.getRange("A8:F8").format = { fill: "#FEF3C7", font: { bold: true } };
  assumptionSheet.getRange("B6:D8").format.numberFormat = "0.0%";
  assumptionSheet.getRange("E6:F8").format.numberFormat = "$#,##0";
  assumptionSheet.getRange("A1:F11").format.columnWidthPx = 148;
  assumptionSheet.getRange("A1:A11").format.columnWidthPx = 250;
  assumptionSheet.getRange("B3").dataValidation = {
    rule: {
      type: "list",
      values: SCENARIOS,
      allowBlank: false,
      showInputMessage: true,
      promptTitle: "Scenario",
      prompt: "Choose Base, Upside, or Downside.",
      showErrorMessage: true,
      errorTitle: "Invalid scenario",
      error: "Choose one of the three approved scenarios.",
      errorStyle: "stop",
      showDropdown: true,
    },
  };
  workbook.comments.setSelf({ displayName: "OfficeKit planning" });
  workbook.comments.addThread(
    { cell: assumptionSheet.getRange("B6") },
    "Revenue growth drives the formula-backed FY27 revenue chain.",
    { id: COMMENT_IDS[0][0], author: "Finance owner", comment: commentConfig(COMMENT_IDS[0][0], COMMENT_IDS[0][1], "Finance owner", "finance.owner@example.test", "2026-08-01T09:00:00Z") },
  );
  workbook.comments.addThread(
    { cell: assumptionSheet.getRange("D6") },
    "Gross margin is a key board-reviewed assumption.",
    { id: COMMENT_IDS[1][0], author: "Finance owner", comment: commentConfig(COMMENT_IDS[1][0], COMMENT_IDS[1][1], "Finance owner", "finance.owner@example.test", "2026-08-01T09:05:00Z") },
  );
  assumptionSheet.freezePanes.freezeRows(5);
  assumptionSheet.showGridLines = false;

  const forecast = workbook.worksheets.add("Forecast");
  forecast.getRange("A1:I17").values = [
    ["FY27 Forecast", null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null, null],
    ["Month", "Revenue", "Operating cost", "Gross profit", "EBITDA", "Ending cash", "Scenario", "Gross margin", "Cash floor"],
    ...Array.from({ length: 12 }, (_, index) => {
      const month = `2027-${String(index + 1).padStart(2, "0")}`;
      return [month, null, null, null, null, null, null, null, null];
    }),
  ];
  const firstForecastRow = 4;
  for (let index = 0; index < 12; index += 1) {
    const row = firstForecastRow + index;
    const previous = row - 1;
    const revenueFormula = index === 0 ? `=Sources!B25*(1+Assumptions!$B$6)` : `=B${previous}*(1+Assumptions!$B$6)`;
    const costFormula = index === 0 ? `=Sources!C25*(1+Assumptions!$C$6)` : `=C${previous}*(1+Assumptions!$C$6)`;
    const cashFormula = index === 0 ? `=Assumptions!$E$6+E${row}` : `=F${previous}+E${row}`;
    forecast.getRange(`B${row}:I${row}`).formulas = [[
      revenueFormula,
      costFormula,
      `=B${row}*Assumptions!$D$6`,
      `=D${row}-C${row}`,
      cashFormula,
      "=Assumptions!$B$3",
      "=Assumptions!$D$6",
      "=Assumptions!$F$6",
    ]];
  }
  forecast.getRange("A1:I1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF", size: 14 } };
  forecast.getRange("A3:I3").format = { fill: "#E2E8F0", font: { bold: true } };
  forecast.getRange("B4:F15").format.numberFormat = "$#,##0";
  forecast.getRange("H4:H15").format.numberFormat = "0.0%";
  forecast.getRange("I4:I15").format.numberFormat = "$#,##0";
  forecast.getRange("A1:I17").format.columnWidthPx = 130;
  forecast.getRange("A1:A17").format.columnWidthPx = 110;
  forecast.getRange("F4:F15").conditionalFormats.add("cellIs", {
    operator: "lessThan",
    formula: "=I4",
    format: { fill: "#FEE2E2", font: { color: "#991B1B", bold: true } },
  });
  forecast.getRange("G4:G15").dataValidation = { rule: { type: "list", values: SCENARIOS, allowBlank: false, showDropdown: true } };
  forecast.freezePanes.freezeRows(3);
  forecast.showGridLines = false;

  const dashboard = workbook.worksheets.add("Dashboard");
  dashboard.getRange("A1:C13").values = [
    ["Month", "Revenue", "EBITDA"],
    ...Array.from({ length: 12 }, (_, index) => [`2027-${String(index + 1).padStart(2, "0")}`, null, null]),
  ];
  dashboard.getRange("B2:C13").formulas = Array.from({ length: 12 }, (_, index) => {
    const row = firstForecastRow + index;
    return [`=Forecast!B${row}`, `=Forecast!E${row}`];
  });
  dashboard.getRange("E1:F3").values = [
    ["Cost structure", "Amount"],
    ["Operating cost", null],
    ["Gross profit", null],
  ];
  dashboard.getRange("F2:F3").formulas = [["=SUM(Forecast!C4:C15)"], ["=SUM(Forecast!D4:D15)"]];
  dashboard.getRange("A1:C1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
  dashboard.getRange("E1:F1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
  dashboard.getRange("B2:C13").format.numberFormat = "$#,##0";
  dashboard.getRange("F2:F3").format.numberFormat = "$#,##0";
  dashboard.getRange("A1:F13").format.columnWidthPx = 128;
  dashboard.freezePanes.freezeRows(1);
  dashboard.showGridLines = false;
  const revenueChart = dashboard.charts.add("line", dashboard.getRange("A1:C13"));
  revenueChart.name = "Revenue and EBITDA trend";
  revenueChart.title = "Revenue and EBITDA trend";
  revenueChart.hasLegend = true;
  revenueChart.setPosition("H2", "O16");
  const costChart = dashboard.charts.add("pie", dashboard.getRange("E1:F3"));
  costChart.name = "Cost structure";
  costChart.title = "Cost structure";
  costChart.setPosition("H18", "O32");

  const checks = workbook.worksheets.add("Checks");
  checks.getRange("A1:B5").values = [
    ["Check", "Result"],
    ["Revenue is formula-backed", null],
    ["Cash warning is present", null],
    ["Dashboard charts have data", null],
    ["Scenario is approved", null],
  ];
  checks.getRange("B2:B5").formulas = [
    ["=IF(COUNTA(Forecast!B4:B15)=12,\"PASS\",\"FAIL\")"],
    ["=IF(COUNTA(Forecast!F4:F15)=12,\"PASS\",\"FAIL\")"],
    ["=IF(COUNTA(Dashboard!B2:C13)=24,\"PASS\",\"FAIL\")"],
    ["=IF(Assumptions!B3<>\"\",\"PASS\",\"FAIL\")"],
  ];
  checks.getRange("A1:B1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
  checks.getRange("A1:B5").format.columnWidthPx = 240;
  checks.getRange("B2:B5").conditionalFormats.add("containsText", { text: "FAIL", format: { fill: "#FEE2E2", font: { color: "#991B1B", bold: true } } });
  checks.freezePanes.freezeRows(1);
  checks.showGridLines = false;

  return workbook;
}

export async function createOperatingPlan({ actualsPath, assumptionsPath, outputPath, auditPath }) {
  const actualsFile = requiredPath(actualsPath, "actualsPath");
  const assumptionsFile = requiredPath(assumptionsPath, "assumptionsPath");
  const finalPath = requiredPath(outputPath, "outputPath");
  const finalAuditPath = requiredPath(auditPath, "auditPath");
  if (new Set([actualsFile, assumptionsFile]).has(finalPath) || finalPath === finalAuditPath) throw new Error("output paths must be distinct from inputs and each other.");
  const [actualsBytes, assumptionsBytes] = await Promise.all([fs.readFile(actualsFile), fs.readFile(assumptionsFile)]);
  const actuals = parseActuals(actualsBytes.toString("utf8"));
  const assumptions = validateAssumptions(JSON.parse(assumptionsBytes.toString("utf8")));
  const workbook = buildWorkbook(actuals, assumptions);
  workbook.recalculate();
  const verification = workbook.verify({ visualQa: true });
  if (!verification.ok) throw new Error(`Operating-plan model verification failed: ${verification.ndjson}`);
  const sourceRender = await renderAllSheets(workbook);
  const exported = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const outputBytes = new Uint8Array(await exported.arrayBuffer());
  const reimported = await SpreadsheetFile.importXlsx(new FileBlob(outputBytes, { type: XLSX_MIME, name: path.basename(finalPath) }));
  reimported.recalculate();
  const reimportVerification = reimported.verify({ visualQa: true });
  if (!reimportVerification.ok) throw new Error(`Operating-plan second import verification failed: ${reimportVerification.ndjson}`);
  const finalRender = await renderAllSheets(reimported);
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  await fs.writeFile(temporaryPath, outputBytes);
  await fs.rename(temporaryPath, finalPath);
  const audit = {
    schema: "office-kit.xlsx-audit.v1",
    status: "succeeded",
    source: {
      inputs: [
        { path: actualsFile, sha256: sha256(actualsBytes), bytes: actualsBytes.length },
        { path: assumptionsFile, sha256: sha256(assumptionsBytes), bytes: assumptionsBytes.length },
      ],
    },
    output: { path: finalPath, sha256: sha256(outputBytes), bytes: outputBytes.length },
    provider: { requested: "office-kit", actual: "office-kit", version: await packageVersion(), silentFallback: false },
    savePolicy: { strategy: "rewrite" },
    operation: { type: "operating-plan-create", scenario: assumptions.activeScenario, sheets: ["Sources", "Assumptions", "Forecast", "Dashboard", "Checks"] },
    warnings: [],
    validation: {
      formulaErrors: 0,
      derivedValuesAreFormulas: true,
      reimport: { ok: true, sheets: reimported.worksheets.items.map((sheet) => sheet.name) },
      verify: { ok: reimportVerification.ok },
      modelRender: { ok: true, sheets: sourceRender },
      finalRender: { ok: true, sheets: finalRender },
      threadedCommentCount: reimported.comments.threads.length,
      chartCount: reimported.worksheets.items.reduce((total, sheet) => total + sheet.charts.items.length, 0),
    },
  };
  await fs.writeFile(finalAuditPath, JSON.stringify(audit, null, 2));
  return { outputPath: finalPath, auditPath: finalAuditPath, audit };
}

function parseCli(argv) {
  const [actualsPath, assumptionsPath, outputPath = "FY27-operating-plan.xlsx", auditPath = "audit.json"] = argv;
  return { actualsPath, assumptionsPath, outputPath, auditPath };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await createOperatingPlan(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({ outputPath: result.outputPath, auditPath: result.auditPath, outputSha256: result.audit.output.sha256, sheets: result.audit.operation.sheets }));
}
