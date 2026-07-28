#!/usr/bin/env node
/**
 * Run the narrow, read-only cross-page ruled-table workflow.
 *
 * This composes a selected pdfplumber runtime with Poppler evidence.  It does
 * not modify the source PDF, retry through another provider, or treat an
 * arbitrary PDF layout as a table when the ruled-table profile is not proven.
 */

import crypto from "node:crypto";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const scripts = path.resolve(here, "..", "scripts");

function usage() {
  return [
    "usage: node officekit-ruled-cross-page-table-workflow.mjs INPUT.pdf \\",
    "  --table-title 'Regional Revenue' --expected-columns 4 --header-rows 2 --min-pages 3 \\",
    "  --footnote-prefix '*' --json outputs/table.json --csv outputs/table.csv \\",
    "  --audit outputs/audit.json --render-dir tmp/pdfs/table-review",
  ].join("\n");
}

function fail(message) {
  throw new Error(`${message}\n\n${usage()}`);
}

function parseArguments(argv) {
  const positionals = [];
  const flags = new Map();
  for (let index = 0; index < argv.length; index += 1) {
    const value = argv[index];
    if (!value.startsWith("--")) {
      positionals.push(value);
      continue;
    }
    const name = value.slice(2);
    if (!name || flags.has(name)) fail(`duplicate or empty option ${value}`);
    const argument = argv[index + 1];
    if (!argument || argument.startsWith("--")) fail(`option ${value} needs a value`);
    flags.set(name, argument);
    index += 1;
  }
  if (positionals.length !== 1) fail("exactly one input PDF is required");
  for (const name of ["table-title", "expected-columns", "json", "csv", "audit", "render-dir"]) {
    if (!flags.has(name)) fail(`--${name} is required`);
  }
  const integer = (name, defaultValue) => {
    const raw = flags.get(name) ?? String(defaultValue);
    if (!/^\d+$/.test(raw) || Number(raw) < 1) fail(`--${name} must be a positive integer`);
    return Number(raw);
  };
  const expectedColumns = integer("expected-columns", 0);
  if (expectedColumns < 2) fail("--expected-columns must be at least 2");
  return {
    input: positionals[0],
    title: flags.get("table-title"),
    expectedColumns,
    headerRows: integer("header-rows", 2),
    minPages: integer("min-pages", 2),
    footnotePrefix: flags.get("footnote-prefix"),
    json: flags.get("json"),
    csv: flags.get("csv"),
    audit: flags.get("audit"),
    renderDir: flags.get("render-dir"),
  };
}

async function existingPath(target) {
  try { return await fs.lstat(target); } catch (error) { if (error.code === "ENOENT") return null; throw error; }
}

async function requireNewFile(target, label) {
  const resolved = path.resolve(target);
  if (await existingPath(resolved)) throw new Error(`${label} must not overwrite an existing path: ${resolved}`);
  return resolved;
}

async function regularInput(target) {
  const resolved = path.resolve(target);
  const stat = await fs.lstat(resolved);
  if (!stat.isFile() || stat.isSymbolicLink()) throw new Error(`input must be a regular non-symlink file: ${resolved}`);
  return resolved;
}

async function evidence(target) {
  const bytes = await fs.readFile(target);
  return {
    path: path.resolve(target),
    bytes: bytes.length,
    sha256: crypto.createHash("sha256").update(bytes).digest("hex"),
  };
}

function run(command, args, label) {
  const completed = spawnSync(command, args, {
    cwd: process.cwd(),
    encoding: "utf8",
    env: { ...process.env, PYTHONDONTWRITEBYTECODE: "1" },
    maxBuffer: 32 * 1024 * 1024,
  });
  if (completed.error || completed.status !== 0) {
    const detail = [completed.error?.message, completed.stderr, completed.stdout].filter(Boolean).join("\n").trim();
    throw new Error(`${label} failed${detail ? `: ${detail}` : ""}`);
  }
  return completed.stdout;
}

function pngDimensions(bytes, label) {
  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  if (bytes.length < 24 || !bytes.subarray(0, 8).equals(signature) || bytes.subarray(12, 16).toString("ascii") !== "IHDR") {
    throw new Error(`${label} is not a complete PNG raster`);
  }
  const width = bytes.readUInt32BE(16);
  const height = bytes.readUInt32BE(20);
  if (!width || !height) throw new Error(`${label} has invalid PNG dimensions`);
  return { width, height };
}

function overlaySvg(width, height, bbox, title, sourcePng) {
  const [x0, top, x1, bottom] = bbox.map(Number);
  if (![x0, top, x1, bottom].every(Number.isFinite) || x1 <= x0 || bottom <= top) throw new Error("table bbox is invalid for Poppler overlay");
  const safeTitle = String(title).replace(/[&<>]/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" })[character]);
  const labelWidth = Math.min(width - x0, Math.max(140, safeTitle.length * 7));
  const embeddedRaster = sourcePng.toString("base64");
  return `<svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" xmlns="http://www.w3.org/2000/svg"><image href="data:image/png;base64,${embeddedRaster}" width="${width}" height="${height}"/><rect x="${x0}" y="${top}" width="${x1 - x0}" height="${bottom - top}" fill="none" stroke="#e63946" stroke-width="2"/><rect x="${x0}" y="${Math.max(0, top - 18)}" width="${labelWidth}" height="17" fill="#e63946"/><text x="${x0 + 4}" y="${Math.max(12, top - 5)}" fill="white" font-family="Arial" font-size="11">${safeTitle}</text></svg>`;
}

async function writeAtomic(target, value) {
  const parent = path.dirname(target);
  await fs.mkdir(parent, { recursive: true });
  const temporary = path.join(parent, `.${path.basename(target)}.tmp-${process.pid}-${crypto.randomBytes(8).toString("hex")}`);
  try {
    await fs.writeFile(temporary, value, { encoding: "utf8", flag: "wx" });
    await fs.rename(temporary, target);
  } catch (error) {
    await fs.unlink(temporary).catch(() => {});
    throw error;
  }
}

async function renderAndOverlay({ source, result, renderDir, pdftoppm }) {
  if (await existingPath(renderDir)) throw new Error(`render directory must not already exist: ${renderDir}`);
  await fs.mkdir(renderDir, { recursive: true });
  const prefix = path.join(renderDir, "source");
  run(pdftoppm, ["-png", "-r", "144", source, prefix], "Poppler render");
  const entries = await fs.readdir(renderDir);
  const sourcePngs = entries
    .map((entry) => ({ entry, match: entry.match(/^source-(\d+)\.png$/) }))
    .filter((entry) => entry.match)
    .sort((left, right) => Number(left.match[1]) - Number(right.match[1]));
  if (sourcePngs.length !== result.table.pageRange.length) {
    throw new Error(`Poppler rendered ${sourcePngs.length} pages; table extraction selected ${result.table.pageRange.length}`);
  }
  const overlays = [];
  for (const segment of result.table.segments) {
    const sourcePng = path.join(renderDir, `source-${segment.page}.png`);
    const sourcePngBytes = await fs.readFile(sourcePng);
    const { width, height } = pngDimensions(sourcePngBytes, `Poppler page ${segment.page}`);
    const xScale = width / Number(segment.pageSize.width);
    const yScale = height / Number(segment.pageSize.height);
    if (!Number.isFinite(xScale) || !Number.isFinite(yScale) || xScale <= 0 || yScale <= 0) throw new Error(`Poppler page ${segment.page} has invalid scale evidence`);
    const [x0, top, x1, bottom] = segment.tableBBox;
    const scaled = [x0 * xScale, top * yScale, x1 * xScale, bottom * yScale];
    // An SVG embeds the exact Poppler raster before drawing the review box.
    // This keeps the shipped workflow self-contained: `sharp` is an optional
    // package peer and must not become an accidental clean-install dependency.
    const overlay = path.join(renderDir, `table-overlay-${segment.page}.svg`);
    await writeAtomic(overlay, overlaySvg(width, height, scaled, `ruled table: ${result.table.title}`, sourcePngBytes));
    const overlayStat = await fs.stat(overlay);
    if (overlayStat.size < 1_000) throw new Error(`table overlay for page ${segment.page} is unexpectedly small`);
    overlays.push({ page: segment.page, source: sourcePng, overlay, bytes: overlayStat.size, tableBBox: segment.tableBBox });
  }
  return { status: "passed", renderer: "pdftoppm", dpi: 144, pageCount: sourcePngs.length, overlays };
}

async function main() {
  const args = parseArguments(process.argv.slice(2));
  const source = await regularInput(args.input);
  const outputs = await Promise.all([
    requireNewFile(args.json, "JSON output"),
    requireNewFile(args.csv, "CSV output"),
    requireNewFile(args.audit, "audit output"),
  ]);
  const [jsonOutput, csvOutput, auditOutput] = outputs;
  if (new Set([source, jsonOutput, csvOutput, auditOutput]).size !== 4) throw new Error("source and all outputs must be distinct paths");
  const renderDir = path.resolve(args.renderDir);
  const sourceBefore = await evidence(source);
  const providerPython = process.env.OFFICE_KIT_PDF_PROVIDER_PYTHON || "python3";
  const pdftoppm = process.env.OFFICE_KIT_PDF_PDFTOPPM || "pdftoppm";
  const commandTrace = [];
  const invoke = (command, commandArgs, label) => {
    commandTrace.push([command, ...commandArgs]);
    return run(command, commandArgs, label);
  };
  try {
    invoke(providerPython, [path.join(scripts, "pdf_provider.py"), "check", "--provider", "pdfplumber", "--require"], "pdfplumber capability probe");
    invoke(providerPython, [path.join(scripts, "pdf_provider.py"), "plan", "--task", "extract", "--provider", "pdfplumber", "--strategy", "read-only", "--input", source, "--require-provider"], "pdfplumber read-only plan");
    const extractorArguments = [
      path.join(scripts, "extract_ruled_table.py"), source,
      "--table-title", args.title,
      "--expected-columns", String(args.expectedColumns),
      "--header-rows", String(args.headerRows),
      "--min-pages", String(args.minPages),
      "--output", jsonOutput,
      "--csv-output", csvOutput,
    ];
    if (args.footnotePrefix !== undefined) extractorArguments.push("--footnote-prefix", args.footnotePrefix);
    invoke(providerPython, extractorArguments, "ruled-table extraction");
    const result = JSON.parse(await fs.readFile(jsonOutput, "utf8"));
    const visual = await renderAndOverlay({ source, result, renderDir, pdftoppm });
    commandTrace.push([pdftoppm, "-png", "-r", "144", source, path.join(renderDir, "source")]);
    const sourceAfter = await evidence(source);
    if (sourceBefore.sha256 !== sourceAfter.sha256 || sourceBefore.bytes !== sourceAfter.bytes) throw new Error("read-only workflow changed the source PDF");
    const jsonEvidence = await evidence(jsonOutput);
    const csvEvidence = await evidence(csvOutput);
    const audit = {
      schema: "office-kit.pdf-audit.v1",
      status: "succeeded",
      source: sourceAfter,
      output: jsonEvidence,
      outputs: { json: jsonEvidence, csv: csvEvidence },
      provider: { actual: "pdfplumber", version: result.provider.version, silentFallback: false },
      savePolicy: { strategy: "read-only", sourceOverwrite: false, artifactWritten: true },
      preflight: { probeCompleted: true, planCompleted: true, sourceInspectionCompleted: true },
      operation: {
        type: "extract-ruled-cross-page-table",
        profile: result.operation.profile,
        tableTitle: result.operation.tableTitle,
        pageRange: result.table.pageRange,
      },
      validation: {
        sourceIdentity: { sourcePreserved: true, before: sourceBefore, after: sourceAfter },
        ruledTable: result.validation,
        poppler: visual,
        noNarrativeReconstruction: true,
      },
      commands: commandTrace,
    };
    await writeAtomic(auditOutput, `${JSON.stringify(audit, null, 2)}\n`);
    invoke(providerPython, [path.join(scripts, "pdf_audit.py"), "validate", auditOutput, "--source", source, "--artifact", jsonOutput, "--require-operation", "extract-ruled-cross-page-table"], "read-only audit validation");
    console.log(JSON.stringify({ ok: true, source: sourceAfter, output: jsonEvidence, csvOutput: csvEvidence, audit: await evidence(auditOutput), visual, silentFallback: false }, null, 2));
  } catch (error) {
    await Promise.all([jsonOutput, csvOutput, auditOutput].map((target) => fs.unlink(target).catch(() => {})));
    throw error;
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main().catch((error) => {
    console.error(JSON.stringify({ ok: false, error: error.message, silentFallback: false }));
    process.exitCode = 2;
  });
}

export { main as extractRuledCrossPageTable };
