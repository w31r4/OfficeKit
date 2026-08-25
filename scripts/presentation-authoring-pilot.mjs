#!/usr/bin/env node

import assert from "node:assert/strict";
import { spawn, spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import {
  mkdir,
  readdir,
  readFile,
  stat,
  writeFile,
} from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import { scanAgentPolicy } from "./pptx-programmable-import-codex-harness.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const manifestPath = path.join(repoRoot, "evals/presentation-authoring-compiler/pilot.v1.json");
const DEFAULT_TIMEOUT_MS = 30 * 60 * 1000;
const MAX_CAPTURE_BYTES = 64 * 1024 * 1024;
const TASK_ID = /^t_[a-f0-9]{12}$/u;

export async function loadPilotManifest() {
  const manifest = JSON.parse(await readFile(manifestPath, "utf8"));
  assert.equal(manifest.schema, "office-kit/presentation-authoring-pilot/v1");
  assert.equal(manifest.tasks.length, 10);
  assert.deepEqual(Object.keys(manifest.arms).sort(), ["A", "B", "C"]);
  assert.equal(manifest.design.totalRuns, 60);
  return manifest;
}

export function buildPilotMatrix(manifest, { taskId, arm, trial } = {}) {
  const tasks = taskId ? manifest.tasks.filter((task) => task.id === taskId) : manifest.tasks;
  if (!tasks.length) throw new Error(`Unknown pilot task ${taskId}`);
  const arms = arm ? [arm] : ["A", "B", "C"];
  for (const selected of arms) if (!manifest.arms[selected]) throw new Error(`Unknown pilot arm ${selected}`);
  const trials = trial == null ? [1, 2] : [positiveInteger(trial, "trial", 2)];
  const matrix = [];
  for (const task of tasks) {
    for (const repetition of trials) {
      const order = stableArmOrder(task.id, repetition, arms);
      for (const selected of order) matrix.push({ task, arm: selected, trial: repetition, armOrder: order });
    }
  }
  return matrix;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const manifest = await loadPilotManifest();
  const matrix = buildPilotMatrix(manifest, {
    taskId: args.task,
    arm: args.arm,
    trial: args.trial,
  });
  if (args["dry-run"]) {
    process.stdout.write(`${JSON.stringify({
      schema: "office-kit/presentation-authoring-pilot-plan/v1",
      package: manifest.design.package,
      totalRuns: matrix.length,
      runs: matrix.map(({ task, arm, trial, armOrder }) => ({ taskId: task.id, scenario: task.scenario, arm, trial, armOrder })),
    }, null, 2)}\n`);
    return;
  }
  const runRoot = path.resolve(required(args, "run-root"));
  await requireAbsent(runRoot);
  await mkdir(runRoot, { recursive: true });
  const packageInfo = await packCandidate(runRoot);
  const records = [];
  for (const entry of matrix) {
    records.push(await runPilotTrial({
      manifest,
      packageInfo,
      runRoot,
      ...entry,
      timeoutMs: args["timeout-ms"] ? positiveInteger(args["timeout-ms"], "timeout-ms", 24 * 60 * 60 * 1000) : DEFAULT_TIMEOUT_MS,
      codexBin: args.codex || process.env.OFFICEKIT_CODEX_BIN || "codex",
    }));
  }
  const result = {
    schema: "office-kit/presentation-authoring-pilot-runs/v1",
    manifestSha256: sha256(await readFile(manifestPath)),
    package: packageInfo,
    environment: {
      platform: process.platform,
      arch: process.arch,
      node: process.version,
      npm: versionLine(process.platform === "win32" ? "npm.cmd" : "npm", ["--version"]),
    },
    runs: records,
    acceptance: {
      expectedRuns: matrix.length,
      completedRuns: records.length,
      passedRuns: records.filter((record) => record.status === "passed").length,
      status: records.length === matrix.length && records.every((record) => record.status === "passed") ? "passed" : "failed",
    },
  };
  const evidencePath = path.join(runRoot, "runs.v1.json");
  await writeFile(evidencePath, `${JSON.stringify(result, null, 2)}\n`, { flag: "wx" });
  process.stdout.write(`${JSON.stringify({ evidence: evidencePath, acceptance: result.acceptance }, null, 2)}\n`);
  if (result.acceptance.status !== "passed") process.exitCode = 1;
}

export async function runPilotTrial({ manifest, packageInfo, runRoot, task, arm, trial, armOrder, timeoutMs, codexBin }) {
  const runId = `${task.id}/${arm}/${trial}`;
  const trialRoot = path.join(runRoot, "runs", task.id, arm, String(trial));
  const workspace = path.join(trialRoot, "workspace");
  const evidenceRoot = path.join(trialRoot, "evidence");
  await mkdir(evidenceRoot, { recursive: true });
  await mkdir(path.join(workspace, "inputs"), { recursive: true });
  await writeFile(path.join(workspace, "package.json"), `${JSON.stringify({ name: `officekit-pilot-${task.id}-${arm}-${trial}`, private: true }, null, 2)}\n`, { flag: "wx" });
  const install = runRequired(
    process.platform === "win32" ? "npm.cmd" : "npm",
    ["install", "--ignore-scripts", "--no-audit", "--no-fund", packageInfo.tarballPath],
    workspace,
    `${runId}: packed install`,
  );
  await writeFile(path.join(evidenceRoot, "npm-install.stdout.txt"), install.stdout, { flag: "wx" });
  await writeFile(path.join(evidenceRoot, "npm-install.stderr.txt"), install.stderr, { flag: "wx" });
  const officekitBin = path.join(workspace, "node_modules/office-kit/bin/officekit.mjs");
  const installedPackage = JSON.parse(await readFile(path.join(workspace, "node_modules/office-kit/package.json"), "utf8"));
  if (installedPackage.name !== "office-kit" || installedPackage.version !== packageInfo.version) throw new Error(`${runId}: packed identity mismatch`);
  const initialized = runRequired(process.execPath, [officekitBin, "init", "--tools", "agents", "--yes", "--json"], workspace, `${runId}: init`);
  await writeFile(path.join(evidenceRoot, "officekit-init.json"), initialized.stdout, { flag: "wx" });
  const prompt = buildPilotPrompt({ manifest, task, arm, trial, armOrder });
  await writeFile(path.join(evidenceRoot, "prompt.md"), `${prompt}\n`, { flag: "wx" });
  const startedAt = Date.now();
  const codex = await runCodex({ codexBin, workspace, evidenceRoot, prompt, timeoutMs });
  const elapsedMs = Date.now() - startedAt;
  const authoredFiles = await readAuthoredFiles(workspace);
  const policy = scanAgentPolicy({ traceText: codex.traceText, authoredFiles });
  const outputPath = path.join(workspace, "outputs", "result.pptx");
  const failures = [];
  if (codex.status !== 0) failures.push(`codex-exit-${codex.status}`);
  if (!policy.passed) failures.push(...policy.findings.map(({ code }) => `policy-${code}`));
  const taskState = await inspectPilotTask(workspace, arm).catch((error) => ({ passed: false, reason: errorMessage(error) }));
  if (!taskState.passed) failures.push(`task-${taskState.reason || "failed"}`);
  const output = await verifyOutput({ officekitBin, workspace, outputPath, evidenceRoot, authoringPlanPath: taskState.planPath }).catch((error) => ({ passed: false, reason: errorMessage(error) }));
  if (!output.passed) failures.push(`output-${output.reason || "failed"}`);
  const run = {
    schema: "office-kit/presentation-authoring-pilot-run/v1",
    runId,
    taskId: task.id,
    scenario: task.scenario,
    arm,
    trial,
    armOrder,
    freshContext: true,
    packedCleanInstall: true,
    elapsedMs,
    retryCount: 0,
    attempts: 1,
    tokenUsage: codex.tokenUsage,
    status: failures.length === 0 ? "passed" : "failed",
    failures,
    checks: { codex: omitTrace(codex), policy, output, task: taskState },
    evidenceDirectory: path.relative(runRoot, evidenceRoot),
  };
  await writeFile(path.join(evidenceRoot, "run.json"), `${JSON.stringify(run, null, 2)}\n`, { flag: "wx" });
  await writeBlindPacket({ runRoot, run, output });
  return run;
}

export function buildPilotPrompt({ manifest, task, arm, trial, armOrder }) {
  const route = manifest.arms[arm].routeInstruction;
  const compilerGuardrails = arm === "C"
    ? " For this experimental authoring-compiler route, make the design grammar explicit before composing: set minimumBodyFontSize to at least 22 and minimumCaptionFontSize to at least 20 in the model's font-size units, and never shrink text to make it fit. Shorten, split, or restructure content instead. Across a deck of six or more pages, use at least four distinct composition silhouettes; repair repeatedComposition, densityRhythmJump, and cardWallPattern warnings unless the plan records a concrete reason. Treat those warnings as actions: the final plan-bound review must have no unrecorded design warnings. If a warning is intentional, add designGrammar.intentionalWarnings with its exact type, affected pageIds, and a concrete reason before commit; a prose invariant is not enough. Give each page one dominant reading anchor, use clear foreground/background contrast, and do not let one pale card surface or dark canvas become the default treatment for every page. For self-directed work, keep palette.strict false unless the user supplied a closed brand palette; strict palettes are for authoritative design systems and must enumerate every allowed emitted color, including theme defaults. Before export, run presentation.validateLayout() on the in-memory deck and resolve its issues while preserving the readability floor; then reimport and run the independent review. Encode every quantitative claim as a meaningful relationship (chart, axis, connector, direct label, or spatial comparison) rather than a decorative list. Prefer direct labels and meaningful diagrams over tiny legends or decorative boxes. Favor fewer, stronger elements: use plain text mainly for thesis or section pages; turn comparisons, causality, timelines, and quantities into a diagram, chart, or spatial relationship, and omit supporting labels that are not readable at contact-sheet scale. Compose nodes are declarative: capture the materialized elements returned by `slide.compose()` for later edits or connector targets; do not pass a compose node itself to a connector. For directed connectors, `head` is the arrow at the `from`/start endpoint and `tail` is the arrow at the `to`/end endpoint; use `tail` for a forward arrow. New connectors are behind shapes by default, so call `connector.bringToFront()` when a background panel or other shape would cover the route."
    : "";
  return `You are a fresh OfficeKit Agent context in a blinded authoring pilot.\n\nTask: ${task.goal}\nScenario: ${task.scenario}\nTrial: ${trial}\nArm route: ${route}\nArm order token: ${armOrder.join(",")}\n\nUse only the installed public OfficeKit Skills and the office-kit package. Work from the sentence above, create a useful multi-page PPTX, and keep the task durable: use officekit repl, commit a reviewed draft, start a fresh REPL process to continue one local edit, review again, and publish outputs/result.pptx. The deck must be created with typed OfficeKit primitives and reimported before delivery. Run an independent final review against the self-directed candidate without a source baseline; fix every semantic/layout error such as text overflow or unexpected overlap before commit.${compilerGuardrails} Use visualReview: unavailable when no visual tool exists; never claim visual completion. Use public Help and shipped examples to discover APIs; do not read \`node_modules/office-kit/src\` or other package internals. Do not install optional renderers or call Playwright/Chromium in the packed portable lane; use the public render/verify path and report visualReview: unavailable when no visual capability is present. Do not read .office-kit/tasks directly: treat it as a private store, and never use shell or node filesystem tools such as cat, sed, ls, file, find, or rg on task manifests, plans, candidates, revisions, sessions, or evidence paths; use officekit tasks/repl responses and public package APIs only. Do not use @oai/artifact-tool, Python, HTML/PPTD, raw OOXML, ZIP/XML patching, another writer, or a silent fallback. Do not ask a question unless a missing answer changes the audience or conclusion. Finish only after the published output exists.`;
}

async function verifyOutput({ officekitBin, workspace, outputPath, evidenceRoot, authoringPlanPath }) {
  const descriptor = await stat(outputPath).catch(() => null);
  if (!descriptor?.isFile()) return { passed: false, reason: "missing-result-pptx" };
  const verifier = path.join(evidenceRoot, "verify.mjs");
  const verifierSource = [
    'import { readFile } from "node:fs/promises";',
    'import { FileBlob, PresentationFile, reviewArtifact } from "office-kit";',
    'const file = process.argv[2];',
    'const planPath = process.argv[3] || null;',
    'const bytes = await readFile(file);',
    'const blob = new FileBlob(bytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });',
    'const presentation = await PresentationFile.importPptx(blob);',
    'const authoringPlan = planPath ? JSON.parse(await readFile(planPath, "utf8")) : undefined;',
    'const review = await reviewArtifact(blob, { layout: false, visualReview: "unavailable", ...(authoringPlan ? { authoringPlan } : {}) });',
    'const designWarnings = (review.design?.issues || []).filter((issue) => issue.severity === "warning").map((issue) => issue.type);',
    'const intentionalWarnings = (authoringPlan?.design?.designGrammar?.intentionalWarnings || []).filter((entry) => entry && typeof entry.type === "string" && typeof entry.reason === "string" && entry.reason.trim() && (Array.isArray(entry.pageIds) || entry.pageIds == null)).map((entry) => entry.type);',
    'const unrecordedWarnings = [...new Set(designWarnings)].filter((type) => !intentionalWarnings.includes(type));',
    'console.log(JSON.stringify({ slides: presentation.slides.items.length, reviewVerdict: review.verdict, visualReview: review.visualReview, sha256: presentation.source?.packageSha256 || null, designWarnings, intentionalWarnings, unrecordedWarnings }));',
    'if (authoringPlan && unrecordedWarnings.length > 0) {',
    '  throw new Error("unresolved-design-warnings: " + unrecordedWarnings.join(","));',
    '}',
    '',
  ].join("\n");
  await writeFile(verifier, verifierSource, { flag: "wx" });
  const verifierArgs = [officekitBin, "run", verifier, "--", outputPath];
  if (authoringPlanPath) verifierArgs.push(authoringPlanPath);
  const result = runRequired(process.execPath, verifierArgs, workspace, "packed public presentation verifier");
  const parsed = JSON.parse(result.stdout.trim().split(/\r?\n/u).filter(Boolean).at(-1));
  if (!Number.isInteger(parsed.slides) || parsed.slides < 1) return { passed: false, reason: "no-slides" };
  if (!["passed", "passed-with-limitations"].includes(parsed.reviewVerdict)) return { passed: false, reason: `review-${parsed.reviewVerdict}` };
  return { passed: true, slides: parsed.slides, reviewVerdict: parsed.reviewVerdict, visualReview: parsed.visualReview, designWarnings: parsed.designWarnings };
}

async function inspectPilotTask(workspace, arm) {
  const store = path.join(workspace, ".office-kit", "tasks");
  const entries = await readdir(store, { withFileTypes: true });
  const taskIds = entries.filter((entry) => entry.isDirectory() && TASK_ID.test(entry.name)).map((entry) => entry.name);
  if (taskIds.length !== 1) return { passed: false, reason: `expected-one-task-${taskIds.length}` };
  const taskRoot = path.join(store, taskIds[0]);
  const manifest = JSON.parse(await readFile(path.join(taskRoot, "task.json"), "utf8"));
  const commits = Array.isArray(manifest.commits) ? manifest.commits : [];
  const publications = Array.isArray(manifest.publications) ? manifest.publications : [];
  if (commits.length < 1 || publications.length < 1 || manifest.head?.commitId !== commits.at(-1)?.id) {
    return { passed: false, reason: "task-not-reviewed-and-published", commits: commits.length, publications: publications.length };
  }
  if (arm === "C" && !manifest.plan) return { passed: false, reason: "authoring-plan-missing" };
  return {
    passed: true,
    taskId: taskIds[0],
    commits: commits.length,
    publications: publications.length,
    plan: Boolean(manifest.plan),
    planPath: manifest.plan ? path.join(taskRoot, manifest.plan.path) : null,
  };
}

async function writeBlindPacket({ runRoot, run, output }) {
  const comparisonKey = sha256(`${run.taskId}:${run.trial}`).slice(0, 16);
  const packetId = sha256(`${run.taskId}:${run.arm}:${run.trial}`).slice(0, 16);
  const packetDirectory = path.join(runRoot, "blind", comparisonKey);
  await mkdir(packetDirectory, { recursive: true });
  const packet = {
    schema: "office-kit/presentation-authoring-blind-packet/v1",
    packetId,
    comparisonKey,
    taskId: run.taskId,
    trial: run.trial,
    artifact: output.passed ? { relativePath: path.join(run.evidenceDirectory, "..", "workspace", "outputs", "result.pptx"), sha256: await sha256File(path.join(runRoot, run.evidenceDirectory, "..", "workspace", "outputs", "result.pptx")) } : null,
    machine: output,
    judge: { status: "unjudged", winner: null, notes: null },
  };
  await writeFile(path.join(packetDirectory, `${packetId}.json`), `${JSON.stringify(packet, null, 2)}\n`, { flag: "wx" });
}

async function runCodex({ codexBin, workspace, evidenceRoot, prompt, timeoutMs }) {
  const tracePath = path.join(evidenceRoot, "codex-trace.jsonl");
  const stderrPath = path.join(evidenceRoot, "codex-stderr.txt");
  const finalPath = path.join(evidenceRoot, "codex-final.txt");
  const args = [
    "exec", "--ephemeral", "--ignore-user-config", "--ignore-rules", "--json",
    "--sandbox", "workspace-write", "--skip-git-repo-check", "-C", workspace,
    "--config", "features.plugins=false", "--config", "features.apps=false",
    "--config", "features.multi_agent=false", "--config", "features.tool_search=false",
    "--config", "plugins={}", "--config", "apps._default.enabled=false",
    "--config", "skills.bundled.enabled=false",
    ...(process.env.OFFICEKIT_PPTX_CODEX_MODEL ? ["--model", process.env.OFFICEKIT_PPTX_CODEX_MODEL] : []),
    "-o", finalPath, "-",
  ];
  const child = spawn(codexBin, args, {
    cwd: workspace,
    detached: process.platform !== "win32",
    env: { ...process.env, PATH: `${path.join(workspace, "node_modules/.bin")}${path.delimiter}${process.env.PATH || ""}` },
    stdio: ["pipe", "pipe", "pipe"],
  });
  let traceText = "";
  let stderr = "";
  child.stdout.setEncoding("utf8");
  child.stderr.setEncoding("utf8");
  child.stdout.on("data", (chunk) => { if (Buffer.byteLength(traceText) + Buffer.byteLength(chunk) <= MAX_CAPTURE_BYTES) traceText += chunk; });
  child.stderr.on("data", (chunk) => { if (Buffer.byteLength(stderr) + Buffer.byteLength(chunk) <= MAX_CAPTURE_BYTES) stderr += chunk; });
  child.stdin.end(prompt);
  let timedOut = false;
  const timer = setTimeout(() => {
    timedOut = true;
    terminateProcessTree(child.pid);
  }, timeoutMs);
  const result = await new Promise((resolve) => {
    child.once("close", (status, signal) => resolve({ status: status ?? 1, signal }));
    child.once("error", (error) => resolve({ status: 127, signal: null, error: error.message }));
  });
  clearTimeout(timer);
  await writeFile(tracePath, traceText, { flag: "wx" });
  await writeFile(stderrPath, stderr, { flag: "wx" });
  const finalBytes = await readFile(finalPath).catch(() => null);
  return {
    status: timedOut ? 124 : result.status,
    signal: result.signal,
    timedOut,
    error: result.error || null,
    traceText,
    tokenUsage: extractTokenUsage(traceText),
    traceSha256: sha256(traceText),
    stderrSha256: sha256(stderr),
    finalSha256: finalBytes ? sha256(finalBytes) : null,
  };
}

function extractTokenUsage(trace) {
  let input = 0;
  let output = 0;
  let observed = false;
  for (const line of String(trace).split(/\r?\n/u)) {
    let event;
    try { event = JSON.parse(line); } catch { continue; }
    const usage = event?.usage || event?.item?.usage || event?.response?.usage;
    if (!usage || typeof usage !== "object") continue;
    const inValue = Number(usage.input_tokens ?? usage.prompt_tokens ?? usage.inputTokens ?? 0);
    const outValue = Number(usage.output_tokens ?? usage.completion_tokens ?? usage.outputTokens ?? 0);
    if (Number.isFinite(inValue) || Number.isFinite(outValue)) {
      input += Number.isFinite(inValue) ? inValue : 0;
      output += Number.isFinite(outValue) ? outValue : 0;
      observed = true;
    }
  }
  return { observed, inputTokens: observed ? input : null, outputTokens: observed ? output : null, totalTokens: observed ? input + output : null };
}

export async function packCandidate(runRoot) {
  const packageRoot = path.join(runRoot, "package");
  await mkdir(packageRoot, { recursive: true });
  const packed = runRequired(process.platform === "win32" ? "npm.cmd" : "npm", ["pack", "--ignore-scripts", "--json", "--pack-destination", packageRoot], repoRoot, "npm pack pilot candidate");
  const record = JSON.parse(packed.stdout.trim())[0];
  const tarballPath = path.join(packageRoot, record.filename);
  const metadata = JSON.parse(await readFile(path.join(repoRoot, "package.json"), "utf8"));
  return { name: metadata.name, version: metadata.version, tarballPath, tarballSha256: await sha256File(tarballPath), packedBytes: record.size, unpackedBytes: record.unpackedSize, totalFiles: record.entryCount };
}

async function readAuthoredFiles(workspace) {
  const output = [];
  async function walk(directory) {
    for (const entry of await readdir(directory, { withFileTypes: true })) {
      if (new Set(["node_modules", ".agents", ".office-kit", "inputs", "outputs", ".git"]).has(entry.name)) continue;
      const target = path.join(directory, entry.name);
      if (entry.isDirectory()) await walk(target);
      else if (entry.isFile() && /\.(?:mjs|md|json)$/u.test(entry.name)) output.push(await readFile(target, "utf8"));
    }
  }
  await walk(workspace);
  return output;
}

function stableArmOrder(taskId, trial, arms) {
  return [...arms].sort((left, right) => sha256(`${taskId}:${trial}:${left}`).localeCompare(sha256(`${taskId}:${trial}:${right}`)));
}

function omitTrace(codex) {
  const { traceText, ...record } = codex;
  return record;
}

function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

async function sha256File(file) {
  return sha256(await readFile(file));
}

function terminateProcessTree(pid, signal = "SIGTERM") {
  if (!pid) return;
  if (process.platform !== "win32") {
    for (const descendant of descendantProcessIds(pid).reverse()) {
      try { process.kill(descendant, signal); } catch {}
    }
  }
  try { if (process.platform !== "win32") process.kill(-pid, signal); else process.kill(pid, signal); } catch {}
}

function descendantProcessIds(rootPid) {
  if (process.platform === "win32") return [];
  const result = spawnSync("ps", ["-axo", "pid=,ppid="], { encoding: "utf8" });
  if (result.status !== 0) return [];
  const children = new Map();
  for (const line of result.stdout.split(/\r?\n/u)) {
    const match = line.trim().match(/^(\d+)\s+(\d+)$/u);
    if (!match) continue;
    const pid = Number(match[1]);
    const parent = Number(match[2]);
    if (!children.has(parent)) children.set(parent, []);
    children.get(parent).push(pid);
  }
  const output = [];
  const pending = [...(children.get(rootPid) || [])];
  while (pending.length) {
    const pid = pending.pop();
    output.push(pid);
    pending.push(...(children.get(pid) || []));
  }
  return output;
}

function runRequired(command, args, cwd, label) {
  const result = spawnSync(command, args, { cwd, encoding: "utf8", maxBuffer: MAX_CAPTURE_BYTES });
  if (result.status !== 0) throw new Error(`${label} failed (${result.status}): ${(result.stderr || result.stdout || "").trim()}`);
  return { stdout: result.stdout || "", stderr: result.stderr || "" };
}

function versionLine(command, args) {
  const result = spawnSync(command, args, { encoding: "utf8" });
  return String(result.stdout || result.stderr || "unavailable").trim().split(/\r?\n/u)[0];
}

function parseArgs(argv) {
  const result = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) throw new Error(`Unexpected argument ${token}`);
    const name = token.slice(2);
    if (name === "dry-run") result[name] = true;
    else if (argv[index + 1] && !argv[index + 1].startsWith("--")) result[name] = argv[++index];
    else throw new Error(`Missing value for ${token}`);
  }
  return result;
}

function required(args, name) {
  if (!args[name]) throw new Error(`Missing --${name}`);
  return args[name];
}

function positiveInteger(value, name, maximum = 10) {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 1 || parsed > maximum) throw new Error(`${name} must be an integer from 1 through ${maximum}`);
  return parsed;
}

async function requireAbsent(target) {
  try { await stat(target); }
  catch (error) { if (error.code === "ENOENT") return; throw error; }
  throw new Error(`Evidence root already exists: ${target}`);
}

function errorMessage(error) {
  return error instanceof Error ? error.message : String(error);
}

if (import.meta.url === pathToFileURL(process.argv[1] || "").href) {
  await main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
