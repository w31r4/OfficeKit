#!/usr/bin/env node

import crypto from "node:crypto";
import fs from "node:fs/promises";
import { spawn, spawnSync } from "node:child_process";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";

import { FileBlob, PresentationFile } from "../src/index.mjs";
import { TEMPLATE_GENERATION_SOURCES } from "./pptx-template-generation.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const DEFAULT_ASSETS_DIR = "/Users/zfang/Downloads/飞书20260814-175228";
const DEFAULT_TIMEOUT_MS = 20 * 60 * 1000;
const TOPOLOGY_PARTS = new Set([
  "[Content_Types].xml",
  "ppt/presentation.xml",
  "ppt/_rels/presentation.xml.rels",
]);

function sha256(value) {
  return crypto.createHash("sha256").update(value).digest("hex");
}

function run(command, args, cwd, label = `${command} ${args.join(" ")}`) {
  const result = spawnSync(command, args, {
    cwd,
    encoding: "utf8",
    stdio: ["ignore", "pipe", "pipe"],
    maxBuffer: 64 * 1024 * 1024,
  });
  if (result.status !== 0) {
    throw new Error(`${label} failed (${result.status})\n${result.stdout || ""}\n${result.stderr || ""}`);
  }
  return result.stdout || "";
}

async function runCodex(workspace, evaluator, phase, prompt, timeoutMs) {
  const phaseRoot = path.join(evaluator, "phases", phase);
  await fs.mkdir(phaseRoot, { recursive: true });
  const args = [
    "exec", "--ephemeral", "--ignore-user-config", "--ignore-rules", "--json",
    "--sandbox", "workspace-write", "--skip-git-repo-check", "-C", workspace,
    "--config", "features.plugins=false",
    "--config", "features.apps=false",
    "--config", "features.multi_agent=false",
    "--config", "features.tool_search=false",
    "--config", "plugins={}",
    "--config", "apps._default.enabled=false",
    "--config", "skills.bundled.enabled=false",
    ...(process.env.OFFICE_KIT_TEMPLATE_AGENT_MODEL ? ["--model", process.env.OFFICE_KIT_TEMPLATE_AGENT_MODEL] : []),
    "-o", path.join(phaseRoot, "final.txt"), "-",
  ];
  const child = spawn(process.env.OFFICE_KIT_CODEX_BIN || "codex", args, {
    cwd: workspace,
    detached: process.platform !== "win32",
    env: {
      ...process.env,
      PATH: `${path.join(workspace, "node_modules/.bin")}${path.delimiter}${process.env.PATH || ""}`,
      PYTHONDONTWRITEBYTECODE: "1",
    },
    stdio: ["pipe", "pipe", "pipe"],
  });
  let stdout = "";
  let stderr = "";
  child.stdout.setEncoding("utf8");
  child.stderr.setEncoding("utf8");
  child.stdout.on("data", (chunk) => { stdout += chunk; });
  child.stderr.on("data", (chunk) => { stderr += chunk; });
  child.stdin.end(prompt);
  let timedOut = false;
  const timer = setTimeout(() => {
    timedOut = true;
    try {
      if (process.platform !== "win32") process.kill(-child.pid, "SIGTERM");
      else child.kill("SIGTERM");
    } catch {}
    // A descendant can retain the inherited JSON pipes after the process group
    // is terminated. Close our ends and resolve the phase rather than waiting
    // forever for a close event that the OS will never deliver.
    setTimeout(() => {
      try { child.stdout.destroy(); } catch {}
      try { child.stderr.destroy(); } catch {}
    }, 5000).unref();
  }, timeoutMs);
  const result = await new Promise((resolve) => {
    let settled = false;
    const finish = (value) => {
      if (settled) return;
      settled = true;
      resolve(value);
    };
    child.once("close", (status, signal) => finish({ status, signal }));
    setTimeout(() => {
      if (timedOut) finish({ status: null, signal: "SIGTERM" });
    }, timeoutMs + 7000).unref();
  });
  clearTimeout(timer);
  const status = timedOut ? 124 : result.status;
  await fs.writeFile(path.join(phaseRoot, "trace.jsonl"), stdout, "utf8");
  await fs.writeFile(path.join(phaseRoot, "stderr.txt"), stderr, "utf8");
  await fs.writeFile(path.join(phaseRoot, "exit.json"), JSON.stringify({ status, signal: result.signal, timedOut, timeoutMs }, null, 2));
  return status;
}

function slideText(slide) {
  return (slide?.shapes?.items || [])
    .flatMap((shape) => (shape.text?.paragraphs || []).flatMap((paragraph) => (paragraph.runs || []).map((run) => String(run.text ?? ""))))
    .join("\u241f");
}

async function packageOracle(sourceBytes, outputBytes) {
  const [sourceZip, outputZip] = await Promise.all([JSZip.loadAsync(sourceBytes), JSZip.loadAsync(outputBytes)]);
  const sourceNames = Object.keys(sourceZip.files).filter((name) => !sourceZip.files[name].dir);
  const changed = [];
  const missing = [];
  for (const name of sourceNames) {
    if (TOPOLOGY_PARTS.has(name) || /^ppt\/slides\/slide\d+(?:\.xml|\.rels)$/u.test(name)) continue;
    const outputEntry = outputZip.files[name];
    if (!outputEntry) {
      missing.push(name);
      continue;
    }
    const [before, after] = await Promise.all([sourceZip.files[name].async("nodebuffer"), outputEntry.async("nodebuffer")]);
    if (!before.equals(after)) changed.push(name);
  }
  return { nonTargetPartsByteIdentical: changed.length === 0 && missing.length === 0, changed, missing };
}

function phasePrompts(sourceId, { mode = "single", pageCount = 1 } = {}) {
  const multiPage = mode === "multi";
  const briefs = Array.from({ length: pageCount }, (_, index) => ({
    role: index === 0 ? "title" : index === pageCount - 1 ? "decision" : "content",
    title: `Agent page ${index + 1}`,
    body: [`Evidence ${index + 1}`, `Decision ${index + 1}`],
  }));
  return {
    plan: `You are the planning phase of an isolated OfficeKit presentation task for ${sourceId}. Read only the installed Presentations Skill and its template-conditioned-generation reference. Use the public office-kit package in a short non-interactive task (run it with officekit run) to import inputs/source.pptx, call designProfile({ maxItems: 64 }), then call planTemplateGeneration({ slides: ${JSON.stringify(briefs)} }). Do not edit or export the presentation. Write outputs/plan.json containing the exact returned plan, sourceSha256, and the requested page briefs. The plan must be ready and contain ${pageCount} pages when the source exposes enough bounded clone-safe slides; otherwise preserve its blocked entries and stop. Do not use @oai/artifact-tool, Python, HTML/PPTD, raw OOXML, XPath, or another writer. Finish after plan.json exists.`,
    author: multiPage
      ? `You are the authoring phase of an isolated OfficeKit presentation task for ${sourceId}. Read outputs/plan.json and use only the public office-kit package. The file is the exact returned plan object; its pages are in plan.pages and each page exposes sourceSlideOrdinal, sourceSlideId, frameTarget, and targetRunText. Write task.mjs and run it exactly once, non-interactively, with officekit run task.mjs. Keep inputs/source.pptx immutable and produce exactly ${pageCount} new content slides from the plan; do not call slides.add, create new shapes, invent a palette, or rebuild a slide. Use the plan's source slide ordinals and frame targets. Because one source origin can have only one pending clone, duplicate at most one planned page per source ordinal in each round, export/reimport between rounds, and resolve the original source by ordinal plus the count of already inserted lower-ordinal clones. After all clones are staged, export/reimport once, locate each clone by source ordinal plus occurrence, and apply exactly one bounded edit on the target: for frameTarget.kind === "shape-run", find the exact existing run text and call shape.text.replace(run.text, replacement); for frameTarget.kind === "svg-text", use slides.items[cloneIndex].images.items[frameTarget.imageIndex].editSvgText(frameTarget.nodeId, { expectedHash: frameTarget.expectedHash, value: replacement }). Never concatenate runs or rewrite an image. Use the plan's title (or first body item) as the replacement value, retaining inherited fonts, geometry, spacing, assets, placeholders, and opaque descendants. Export outputs/generated.pptx and stop once it exists; the outer evaluator performs the reimport and package oracle. Do not launch a second OfficeKit process after output exists. Do not call verify({ visualQa: true }), render, auditAccessibility, or review in this phase. Do not start officekit repl, wait for stdin, use @oai/artifact-tool, Python, HTML/PPTD, raw OOXML, XPath, or another writer. Do not write audit.json; finish when generated.pptx exists.`
      : `You are the authoring phase of an isolated OfficeKit presentation task for ${sourceId}. Read outputs/plan.json and use only the public office-kit package. Write task.mjs and run it exactly once, non-interactively, with officekit run task.mjs. Keep inputs/source.pptx immutable. For this black-box trial, author exactly one new content slide from the first planned clone-safe source slide: do not call slides.add, do not create new shapes, and do not invent a new palette. Use presentation.slides.items[index] (not getItemAt), duplicate one planned clone-safe source slide through one export/reimport boundary, and make exactly one bounded edit on that clone while retaining inherited fonts, geometry, spacing, assets, and opaque descendants. For text, inspect the cloned shape's text.paragraphs and runs and pass the exact text of one existing run to shape.text.replace; never pass a concatenated shape/paragraph value, and never edit across a run or paragraph boundary. Prefer an existing image.svgTextCapability leaf with image.editSvgText when no single run is available; if no bounded leaf exists, stop with a clear fail-closed error instead of retrying. Export outputs/generated.pptx, then stop once the file exists; the outer evaluator performs the reimport and source/package oracle. Do not launch a second OfficeKit process if the first command has produced the output. Do not call verify({ visualQa: true }), render, auditAccessibility, or any review in this phase; the next phase owns review. Do not start officekit repl, wait for stdin, use @oai/artifact-tool, Python, HTML/PPTD, raw OOXML, XPath, or another writer. Do not write audit.json yet; finish when generated.pptx exists.`,
    review: `You are the review phase of an isolated OfficeKit presentation task for ${sourceId}. Read outputs/plan.json and outputs/generated.pptx. Use exactly one short non-interactive OfficeKit task (run it with officekit run) to reimport the output, run the static PresentationFile.inspectPptx(outputBlob, { includeText: false, maxChars: 12000 }) and the instance presentation.validateLayout({ maxChars: 12000 }), and compare the source hash and non-target OPC parts against inputs/source.pptx. Use Node APIs (FileBlob, PresentationFile, JSZip, fs) for the comparison; do not use shell unzip, awk, or a shell variable named status. Keep report strings bounded. Write outputs/audit.json with sourceSha256, outputSha256, sourceSlides, outputSlides, profileSummary, frameMap, reimport, sourceProtected, packageOracle, review, and visualReview. Set sourceProtected to an object with expected/before/after/unchanged, and set visualReview to unavailable when no renderer or visual capability is available. The audit must be written even when inherited layout warnings exist; record those warnings in review rather than retrying indefinitely. Never call structural checks an aesthetic pass. Do not modify inputs/source.pptx and do not use @oai/artifact-tool, Python, HTML/PPTD, raw OOXML, XPath, or another writer. Finish immediately after audit.json exists.`,
  };
}

async function prepareTrial({ repoRoot, assetsDir, definition, packPath, runRoot, trialNumber, mode = "single", pageCount = 1 }) {
  const trialRoot = path.join(runRoot, `${definition.id}-trial-${trialNumber}`);
  const workspace = path.join(trialRoot, "workspace");
  const evaluator = path.join(trialRoot, "evaluator");
  await fs.mkdir(path.join(workspace, "inputs"), { recursive: true });
  await fs.mkdir(path.join(workspace, "outputs"), { recursive: true });
  await fs.mkdir(evaluator, { recursive: true });
  const sourcePath = path.join(assetsDir, definition.fileName);
  const sourceBytes = await fs.readFile(sourcePath);
  const sourceSha256 = sha256(sourceBytes);
  if (sourceSha256 !== definition.sourceSha256) throw new Error(`${definition.id}: source SHA-256 mismatch: ${sourceSha256}`);
  await fs.copyFile(sourcePath, path.join(workspace, "inputs/source.pptx"));
  await fs.writeFile(path.join(workspace, "package.json"), JSON.stringify({ name: `officekit-template-blackbox-${definition.id}`, private: true }, null, 2));
  run(process.platform === "win32" ? "npm.cmd" : "npm", ["install", "--ignore-scripts", "--no-audit", "--no-fund", packPath], workspace, `${definition.id}: npm install`);
  const officekit = path.join(workspace, "node_modules/office-kit/bin/officekit.mjs");
  run(process.execPath, [officekit, "init", "--tools", "agents", "--yes", "--json"], workspace, `${definition.id}: officekit init`);
  for (const skill of ["documents", "spreadsheets", "excel-live-control", "powerpoint-live-control", "pdf", "template-creator"]) {
    await fs.rm(path.join(workspace, ".agents/skills", skill), { recursive: true, force: true });
  }
  const prompts = phasePrompts(definition.id, { mode, pageCount });
  await fs.writeFile(path.join(workspace, "PROMPT.md"), `${prompts.plan}\n`, "utf8");
  await fs.writeFile(path.join(evaluator, "prompt.sha256"), `${sha256(Buffer.from(Object.values(prompts).join("\n"), "utf8"))}\n`, "utf8");
  const timeoutMs = Number(process.env.OFFICE_KIT_TEMPLATE_AGENT_TIMEOUT_MS || DEFAULT_TIMEOUT_MS);
  const phaseStatuses = [];
  await fs.mkdir(path.join(evaluator, "phases"), { recursive: true });
  for (const phase of ["plan", "author", "review"]) {
    await fs.writeFile(path.join(evaluator, "phases", `${phase}.prompt.md`), `${prompts[phase]}\n`, "utf8");
    const phaseStatus = await runCodex(workspace, evaluator, phase, prompts[phase], timeoutMs);
    phaseStatuses.push({ phase, status: phaseStatus });
    if (phaseStatus !== 0) break;
    if (phase === "plan" && !(await fs.stat(path.join(workspace, "outputs/plan.json")).catch(() => null))) {
      phaseStatuses[phaseStatuses.length - 1].status = 125;
      break;
    }
    if (phase === "author" && !(await fs.stat(path.join(workspace, "outputs/generated.pptx")).catch(() => null))) {
      phaseStatuses[phaseStatuses.length - 1].status = 125;
      break;
    }
  }
  const status = phaseStatuses.every(({ status: phaseStatus }) => phaseStatus === 0) ? 0 : (phaseStatuses.at(-1)?.status ?? 1);
  const outputPath = path.join(workspace, "outputs/generated.pptx");
  const auditPath = path.join(workspace, "outputs/audit.json");
  const outputBytes = await fs.readFile(outputPath).catch(() => null);
  const planText = await fs.readFile(path.join(workspace, "outputs/plan.json"), "utf8").catch(() => "");
  let plan = null;
  try { plan = JSON.parse(planText); } catch {}
  const planResult = plan?.plan && typeof plan.plan === "object" ? plan.plan : plan;
  const auditText = await fs.readFile(auditPath, "utf8").catch(() => "");
  let audit = null;
  try { audit = JSON.parse(auditText); } catch {}
  const source = await PresentationFile.importPptx(new FileBlob(sourceBytes, { type: PPTX_MIME, name: definition.fileName }));
  let output = null;
  let reimportError = null;
  if (outputBytes) {
    try { output = await PresentationFile.importPptx(new FileBlob(outputBytes, { type: PPTX_MIME, name: "generated.pptx" })); } catch (error) { reimportError = error instanceof Error ? error.message : String(error); }
  }
  const sourceSlideTexts = source.slides.items.map(slideText);
  const outputSlideTexts = output?.slides.items.map(slideText) || [];
  const sourceSlidesPreserved = sourceSlideTexts.every((value) => outputSlideTexts.includes(value));
  const packageDiff = outputBytes ? await packageOracle(sourceBytes, outputBytes) : { nonTargetPartsByteIdentical: false, changed: [], missing: [] };
  const auditSourceProtected = audit?.sourceProtected === true || audit?.sourceProtected?.unchanged === true;
  const auditReimportPassed = audit?.reimport?.ok === true || audit?.reimport?.imported === true;
  const profileCaptured = Boolean(output && source.designProfile({ maxItems: 64 }));
  const record = {
    trial: `${definition.id}-trial-${trialNumber}`,
    sourceId: definition.id,
    sourceSha256,
    sourceSlides: source.slides.count,
    outputSha256: outputBytes ? sha256(outputBytes) : null,
    outputBytes: outputBytes?.length ?? 0,
    outputSlides: output?.slides.count ?? null,
    generatedSlides: output ? output.slides.count - source.slides.count : null,
    sourceProtected: sha256(await fs.readFile(path.join(workspace, "inputs/source.pptx"))) === sourceSha256,
    sourceSlidesPreserved,
    packageOracle: packageDiff,
    reimport: { passed: Boolean(output) && !reimportError, error: reimportError },
    profileCaptured,
    planCaptured: Boolean(planResult && planResult.schema === "office-kit/pptx-template-plan/v1" && planResult.status === "ready" && Array.isArray(planResult.pages) && planResult.pages.length >= pageCount),
    plannedPages: Array.isArray(planResult?.pages) ? planResult.pages.length : 0,
    mode,
    agent: { command: process.env.OFFICE_KIT_CODEX_BIN || "codex", exitStatus: status, promptSha256: sha256(Buffer.from(Object.values(prompts).join("\n"), "utf8")), phases: phaseStatuses },
    audit: audit ? {
      hasFrameMap: Array.isArray(audit.frameMap) && audit.frameMap.length >= 3,
      hasProfileSummary: Boolean((audit.profileSummary && typeof audit.profileSummary === "object") || profileCaptured),
      sourceProtected: auditSourceProtected,
      reimportPassed: auditReimportPassed,
      reimport: audit.reimport ?? null,
      visualReview: audit.visualReview ?? audit.review?.visualReview ?? null,
      review: audit.review ?? null,
    } : null,
  };
  await fs.writeFile(path.join(evaluator, "summary.json"), `${JSON.stringify(record, null, 2)}\n`, "utf8");
  return record;
}

function parseArgs(argv) {
  const result = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (token.startsWith("--")) result[token.slice(2)] = argv[index + 1] && !argv[index + 1].startsWith("--") ? argv[++index] : true;
  }
  return result;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const assetsDir = path.resolve(args["assets-dir"] || process.env.OFFICE_KIT_TEMPLATE_ASSETS_DIR || DEFAULT_ASSETS_DIR);
  const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
  const runRoot = path.resolve(args["run-root"] || await fs.mkdtemp(path.join(os.tmpdir(), "officekit-template-agent-blackbox-")));
  const mode = args.mode === "multi" ? "multi" : "single";
  const pageCount = Number(args.slides || (mode === "multi" ? 10 : 1));
  if (!Number.isInteger(pageCount) || pageCount < 1 || pageCount > 32) throw new RangeError("--slides must be an integer from 1 through 32");
  await fs.mkdir(runRoot, { recursive: true });
  const packDir = path.join(runRoot, "pack");
  await fs.mkdir(packDir, { recursive: true });
  const packJson = JSON.parse(run(process.platform === "win32" ? "npm.cmd" : "npm", ["pack", "--json", "--ignore-scripts", "--pack-destination", packDir], repoRoot, "npm pack"));
  const packPath = path.resolve(packDir, packJson[0].filename);
  const definitions = args.source ? TEMPLATE_GENERATION_SOURCES.filter((item) => item.id === args.source) : TEMPLATE_GENERATION_SOURCES;
  if (!definitions.length) throw new Error(`unknown source ${args.source}`);
  const trials = [];
  for (const [index, definition] of definitions.entries()) trials.push(await prepareTrial({ repoRoot, assetsDir, definition, packPath, runRoot, trialNumber: index + 1, mode, pageCount }));
  const evidence = {
    schema: "office-kit/pptx-template-conditioned-generation-agent-blackbox/v1",
    generatedAt: new Date().toISOString(),
    package: { version: JSON.parse(await fs.readFile(path.join(repoRoot, "package.json"), "utf8")).version, sha256: sha256(await fs.readFile(packPath)) },
    environment: { platform: process.platform, arch: process.arch, node: process.version, codex: process.env.OFFICE_KIT_CODEX_BIN || "codex" },
    mode,
    pageCount,
    trials,
    acceptance: {
      required: definitions.length,
      completed: trials.filter((trial) => trial.agent.exitStatus === 0).length,
      allReimported: trials.every((trial) => trial.reimport.passed),
      allSourceProtected: trials.every((trial) => trial.sourceProtected && trial.sourceSlidesPreserved),
      allPackageNonTargetPreserved: trials.every((trial) => trial.packageOracle.nonTargetPartsByteIdentical),
      allAuditsPresent: trials.every((trial) => trial.audit?.hasFrameMap && trial.audit?.hasProfileSummary && trial.audit?.sourceProtected && trial.audit?.reimportPassed === true),
      status: trials.length === definitions.length && trials.every((trial) => trial.agent.exitStatus === 0 && trial.reimport.passed && trial.sourceProtected && trial.sourceSlidesPreserved && trial.packageOracle.nonTargetPartsByteIdentical && trial.audit?.hasFrameMap && trial.audit?.hasProfileSummary && trial.audit?.sourceProtected && trial.audit?.reimportPassed === true && (mode !== "multi" || trial.planCaptured && trial.outputSlides >= pageCount + trial.sourceSlides)) ? "passed" : "blocked",
    },
  };
  const evidencePath = path.resolve(args.output || path.join(repoRoot, "evals/pptx-generation/agent-blackbox.v1.json"));
  await fs.mkdir(path.dirname(evidencePath), { recursive: true });
  await fs.writeFile(evidencePath, `${JSON.stringify(evidence, null, 2)}\n`, "utf8");
  console.log(JSON.stringify({ evidencePath, acceptance: evidence.acceptance, trials: trials.map(({ trial, outputSlides, generatedSlides, reimport, sourceProtected, sourceSlidesPreserved, packageOracle, audit }) => ({ trial, outputSlides, generatedSlides, reimport, sourceProtected, sourceSlidesPreserved, nonTargetPartsByteIdentical: packageOracle.nonTargetPartsByteIdentical, audit })) }, null, 2));
  if (evidence.acceptance.status !== "passed") process.exitCode = 1;
}

if (import.meta.url === `file://${process.argv[1]}`) await main();
