import crypto from "node:crypto";
import { constants as fsConstants } from "node:fs";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";

import { FileBlob, SpreadsheetFile } from "office-kit";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const FIXTURE = {
  assumptionsSheetName: "Assumptions",
  assumptionAddress: "B4",
  originalAssumption: 0.065,
  replacementAssumption: 0.07,
  dashboardSheetName: "Dashboard",
  chartName: "enterprise-line-chart",
  originalChartTitle: "Revenue Outlook",
  replacementChartTitle: "Updated Revenue Outlook",
  dataSheetName: "Data",
  pivotSheetName: "Summary",
  pivotName: "Enterprise Revenue",
  dynamicArrayAddress: "G2:G4",
  connectionId: 9,
  connectionName: "Enterprise warehouse",
  customPowerQueryPath: "customXml/itemEnterprisePowerQuery.xml",
  slicerPath: "xl/slicers/slicer1.xml",
  slicerCachePath: "xl/slicerCaches/slicerCache1.xml",
  comboChartPath: "xl/drawings/charts/chart2.xml",
};
const require = createRequire(import.meta.url);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function sameJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty string.`);
  return value.trim();
}

async function packageVersion() {
  const entry = require.resolve("office-kit");
  const packagePath = path.join(path.dirname(path.dirname(entry)), "package.json");
  return JSON.parse(await fs.readFile(packagePath, "utf8")).version;
}

async function assertNewFile(filePath, label) {
  try {
    await fs.access(filePath);
  } catch (error) {
    if (error?.code === "ENOENT") return;
    throw error;
  }
  throw new Error(`${label} already exists; refusing to overwrite it.`);
}

async function publishNoOverwrite(temporaryPath, finalPath, label) {
  try {
    await fs.copyFile(temporaryPath, finalPath, fsConstants.COPYFILE_EXCL);
  } catch (error) {
    if (error?.code === "EEXIST") throw new Error(`${label} already exists; refusing to overwrite it.`);
    throw error;
  }
}

function connectionSnapshot(connection) {
  return {
    connectionId: connection.connectionId,
    name: connection.name,
    type: connection.type,
    refreshedVersion: connection.refreshedVersion,
    description: connection.description,
    keepAlive: connection.keepAlive,
    background: connection.background,
    refreshOnLoad: connection.refreshOnLoad,
    saveData: connection.saveData,
    intervalMinutes: connection.intervalMinutes,
  };
}

function workbookProfile(workbook) {
  const data = workbook.worksheets.items.find((sheet) => sheet.name === FIXTURE.dataSheetName);
  const summary = workbook.worksheets.items.find((sheet) => sheet.name === FIXTURE.pivotSheetName);
  const dynamic = data?.store.get("G2");
  return {
    sheets: workbook.worksheets.items.map((sheet) => sheet.name),
    connections: workbook.connections.map(connectionSnapshot),
    sparklineCount: data?.sparklineGroups.items.length || 0,
    pivotNames: summary?.pivotTables.items.map((pivot) => pivot.name) || [],
    dynamicArray: dynamic ? { formula: dynamic.formula, formulaType: dynamic.formulaType, arrayRef: dynamic.arrayRef } : null,
    commentThreads: workbook.comments.threads.length,
  };
}

async function packageProfile(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const partHashes = {};
  for (const partPath of paths) partHashes[partPath] = sha256(await zip.file(partPath).async("uint8array"));
  const text = async (partPath) => zip.file(partPath)?.async("text") || "";
  const [connections, queryTable, metadata, powerQuery, slicer, slicerCache, comboChart, threadedComments, pivotTable, pivotCache] = await Promise.all([
    text("xl/connections.xml"),
    text("xl/queryTables/queryTable1.xml"),
    text("xl/metadata.xml"),
    text(FIXTURE.customPowerQueryPath),
    text(FIXTURE.slicerPath),
    text(FIXTURE.slicerCachePath),
    text(FIXTURE.comboChartPath),
    text("xl/threadedcomments/threadedcomment.xml"),
    text("xl/pivotTables/pivotTable.xml"),
    text("pivotCache/pivotCacheDefinition1.xml"),
  ]);
  return {
    paths,
    partHashes,
    advanced: {
      connections,
      queryTable,
      metadata,
      powerQuery,
      slicer,
      slicerCache,
      comboChart,
      threadedComments,
      pivotTable,
      pivotCache,
    },
  };
}

function packageDiff(source, output) {
  return [...new Set([...source.paths, ...output.paths])].sort().filter((partPath) => source.partHashes[partPath] !== output.partHashes[partPath]);
}

async function renderAllSheets(workbook) {
  const sheets = [];
  for (const sheet of workbook.worksheets.items) {
    const preview = await workbook.render({ sheetName: sheet.name, autoCrop: "all", format: "svg" });
    const svg = await preview.text();
    if (!/<svg\b/i.test(svg)) throw new Error(`Model render for sheet ${sheet.name} did not produce SVG.`);
    sheets.push({ sheet: sheet.name, bytes: preview.bytes.length, renderer: "model-svg" });
  }
  return sheets;
}

/**
 * Change only the enterprise assumption and ordinary line-chart title in a
 * source-bound workbook.  Advanced graphs are inspected and retained as
 * canaries; this workflow does not claim to edit combo charts, slicers,
 * Power Query, connections, PivotTables, dynamic-array metadata, sparklines,
 * or threaded-comment topology.
 */
export async function editXlsxOpaqueEnterprise({ inputPath, outputPath, auditPath }) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || finalAuditPath === sourcePath || finalAuditPath === finalPath) throw new Error("Input, output, and audit paths must be distinct.");
  await Promise.all([assertNewFile(finalPath, "XLSX output"), assertNewFile(finalAuditPath, "Audit output")]);
  const source = await fs.readFile(sourcePath);
  const workbook = await SpreadsheetFile.importXlsx(new FileBlob(source, { type: XLSX_MIME, name: path.basename(sourcePath) }));
  const assumptions = workbook.worksheets.items.find((sheet) => sheet.name === FIXTURE.assumptionsSheetName);
  const dashboard = workbook.worksheets.items.find((sheet) => sheet.name === FIXTURE.dashboardSheetName);
  if (!assumptions || !dashboard) throw new Error("Enterprise source workbook is missing Assumptions or Dashboard sheets.");
  const sourceValue = assumptions.getRange(FIXTURE.assumptionAddress).values[0][0];
  if (Math.abs(Number(sourceValue) - FIXTURE.originalAssumption) > 1e-12) throw new Error(`Expected ${FIXTURE.assumptionsSheetName}!${FIXTURE.assumptionAddress}=${FIXTURE.originalAssumption}; found ${sourceValue}.`);
  const chart = dashboard.charts.items.find((candidate) => candidate.name === FIXTURE.chartName);
  if (!chart || chart.type !== "line" || chart.title !== FIXTURE.originalChartTitle) throw new Error("The ordinary enterprise line chart is missing or does not match the source-bound title precondition.");
  const sourceProfile = workbookProfile(workbook);
  const sourcePackage = await packageProfile(source);
  assumptions.getRange(FIXTURE.assumptionAddress).values = [[FIXTURE.replacementAssumption]];
  chart.title = FIXTURE.replacementChartTitle;
  const editedProfile = workbookProfile(workbook);
  if (editedProfile.connections.length !== sourceProfile.connections.length || !sameJson(editedProfile.pivotNames, sourceProfile.pivotNames) || editedProfile.sparklineCount !== sourceProfile.sparklineCount || !sameJson(editedProfile.dynamicArray, sourceProfile.dynamicArray) || editedProfile.commentThreads !== sourceProfile.commentThreads) {
    throw new Error("The in-memory transaction changed an advanced enterprise object outside the two approved semantic targets.");
  }
  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  try {
    await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
    const exported = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const outputPackage = await packageProfile(output);
    const changedPaths = packageDiff(sourcePackage, outputPackage);
    const expectedChangedPaths = ["xl/drawings/charts/chart1.xml", "xl/worksheets/sheet1.xml"];
    if (!sameJson(changedPaths, expectedChangedPaths)) throw new Error(`Enterprise edit changed unexpected package parts: ${changedPaths.join(", ")}.`);
    const reimported = await SpreadsheetFile.importXlsx(new FileBlob(output, { type: XLSX_MIME, name: path.basename(finalPath) }));
    const outputAssumptions = reimported.worksheets.items.find((sheet) => sheet.name === FIXTURE.assumptionsSheetName);
    const outputDashboard = reimported.worksheets.items.find((sheet) => sheet.name === FIXTURE.dashboardSheetName);
    const outputChart = outputDashboard?.charts.items.find((candidate) => candidate.name === FIXTURE.chartName);
    const outputProfile = workbookProfile(reimported);
    if (Math.abs(Number(outputAssumptions?.getRange(FIXTURE.assumptionAddress).values[0]?.[0]) - FIXTURE.replacementAssumption) > 1e-12 || outputChart?.title !== FIXTURE.replacementChartTitle || !sameJson(outputProfile.connections, sourceProfile.connections) || !sameJson(outputProfile.pivotNames, sourceProfile.pivotNames) || outputProfile.sparklineCount !== sourceProfile.sparklineCount || !sameJson(outputProfile.dynamicArray, sourceProfile.dynamicArray) || outputProfile.commentThreads !== sourceProfile.commentThreads) {
      throw new Error("Second import did not preserve the two edits and every enterprise canary.");
    }
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error(`Workbook verification failed: ${verification.ndjson}`);
    const renders = await renderAllSheets(reimported);
    const sourceAfter = await fs.readFile(sourcePath);
    if (!source.equals(sourceAfter)) throw new Error("The source workbook changed during the transaction; refusing to publish output.");
    const preservedAdvancedParts = Object.keys(sourcePackage.advanced).every((key) => sourcePackage.advanced[key] === outputPackage.advanced[key]);
    if (!preservedAdvancedParts) throw new Error("An advanced enterprise package part changed outside the approved sheet/chart edits.");
    const audit = {
      schema: "office-kit.xlsx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sha256(source), bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "office-kit", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: {
        type: "opaque-enterprise-local-edit",
        assumption: { sheet: FIXTURE.assumptionsSheetName, address: FIXTURE.assumptionAddress, previous: FIXTURE.originalAssumption, value: FIXTURE.replacementAssumption },
        chart: { sheet: FIXTURE.dashboardSheetName, name: FIXTURE.chartName, previousTitle: FIXTURE.originalChartTitle, title: FIXTURE.replacementChartTitle },
        preservationOnly: ["connection", "queryTable", "PivotTable", "sparkline", "dynamic-array-metadata", "threaded-comments", "combo-chart", "slicer", "slicer-cache", "Power Query"],
      },
      warnings: ["The combo chart, slicer/cache, Power Query, connection, QueryTable, PivotTable, sparkline, dynamic-array metadata, and threaded-comment graph are source-owned canaries; this transaction does not author or refresh them."],
      validation: {
        package: { changedPaths, expectedChangedPaths, advancedPartsPreserved: preservedAdvancedParts },
        reimport: { ok: true, advancedObjectsPreserved: true, assumptionUpdated: true, chartTitleUpdated: true },
        verify: { ok: verification.ok },
        modelRender: { ok: true, sheets: renders },
      },
    };
    await fs.writeFile(temporaryAuditPath, `${JSON.stringify(audit, null, 2)}\n`, "utf8");
    await publishNoOverwrite(temporaryPath, finalPath, "XLSX output");
    try {
      await publishNoOverwrite(temporaryAuditPath, finalAuditPath, "Audit output");
    } catch (error) {
      await fs.rm(finalPath, { force: true });
      throw error;
    }
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } finally {
    await Promise.all([fs.rm(temporaryPath, { force: true }), fs.rm(temporaryAuditPath, { force: true })]);
  }
}

function parseCli(argv) {
  const [inputPath, outputPath, auditPath] = argv;
  return { inputPath, outputPath, auditPath };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editXlsxOpaqueEnterprise(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({ outputPath: result.outputPath, auditPath: result.auditPath, outputSha256: result.audit.output.sha256 }));
}
