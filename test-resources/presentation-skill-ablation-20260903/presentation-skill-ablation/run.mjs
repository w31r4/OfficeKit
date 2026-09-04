#!/usr/bin/env node

/*
 * Disposable runner for the presentation Skill ablation study.
 *
 * This file is intentionally kept under evals/ and is not imported by the
 * OfficeKit package. It prepares frozen evidence, runs bounded Codex author
 * or judge sessions, and writes normalized records. It is not a benchmark
 * framework.
 */

import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { access, cp, mkdir, readFile, readdir, rm, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const EXPERIMENT_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)));
const REPO_ROOT = path.resolve(EXPERIMENT_ROOT, "../..");
const CASES_PATH = path.join(EXPERIMENT_ROOT, process.env.PRESENTATION_ABLATION_CASES || "cases.v1.json");
const RUBRIC_PATH = path.join(EXPERIMENT_ROOT, "rubric.v1.json");
const COVERAGE_PATH = path.join(EXPERIMENT_ROOT, "capability-coverage.v1.json");
const FIXTURES_PATH = path.join(EXPERIMENT_ROOT, "fixtures", "index.v1.json");
const COMMON_ROOT = path.join(EXPERIMENT_ROOT, "common");
const ARMS_ROOT = path.join(EXPERIMENT_ROOT, "arms");
const DEFAULT_TIMEOUT_MS = 20 * 60 * 1000;
const MAX_CAPTURE_BYTES = 2 * 1024 * 1024;
const DEFAULT_SEED = 20260902;

const REQUIRED_COMMON = [
  "invariants.md", "what.md", "what-kind.md", "how.md", "style-brief.md",
  "capability-map.md", "references/ppj.md", "references/fonts.md",
  "references/shapes.md", "references/text.md",
  "references/charts-and-tables.md", "references/media-and-layers.md",
  "references/image-sourcing.md", "references/motion.md",
  "references/imported-native-ref.md", "references/review-and-delivery.md",
];
const REQUIRED_CAPABILITIES = new Set([
  "text", "rich-text", "shape", "line-connector", "image-background",
  "mask-opacity", "chart", "table", "group-z-order", "formula", "motion",
  "source-bound", "review",
]);
const FORBIDDEN_CLEAN_ROOM_MARKERS = [
  "kimi-slides", "pptd.md", "pptd format", "kimi source",
];

function fail(message) {
  throw new Error(message);
}

function parseArgs(argv) {
  const command = argv[0] || "help";
  const flags = new Map();
  const positionals = [];
  for (let index = 1; index < argv.length; index += 1) {
    const value = argv[index];
    if (!value.startsWith("--")) {
      positionals.push(value);
      continue;
    }
    const pair = value.slice(2).split("=", 2);
    if (pair.length === 2) flags.set(pair[0], pair[1]);
    else if (argv[index + 1] && !argv[index + 1].startsWith("--")) flags.set(pair[0], argv[++index]);
    else flags.set(pair[0], true);
  }
  return { command, flags, positionals };
}

function flag(flags, name, fallback) {
  return flags.has(name) ? flags.get(name) : fallback;
}

function boolFlag(flags, name) {
  const value = flag(flags, name, false);
  return value === true || value === "true" || value === "1";
}

function positiveInteger(value, label) {
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed < 1) fail(label + " must be a positive integer");
  return parsed;
}

function nowIso() {
  return new Date().toISOString();
}

async function readJson(file) {
  try {
    return JSON.parse(await readFile(file, "utf8"));
  } catch (error) {
    fail(path.relative(REPO_ROOT, file) + ": " + error.message);
  }
}

function digestBytes(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

async function digestFile(file) {
  return digestBytes(await readFile(file));
}

function digestText(value) {
  return digestBytes(Buffer.from(value, "utf8"));
}

async function exists(file) {
  try {
    await access(file);
    return true;
  } catch {
    return false;
  }
}

async function ensureDir(directory) {
  await mkdir(directory, { recursive: true });
}

async function writeJson(file, value) {
  await ensureDir(path.dirname(file));
  await writeFile(file, JSON.stringify(value, null, 2) + "\n", "utf8");
}

function seededRandom(seed) {
  let state = seed >>> 0;
  return () => {
    state += 0x6d2b79f5;
    let value = state;
    value = Math.imul(value ^ (value >>> 15), value | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return ((value ^ (value >>> 14)) >>> 0) / 4294967296;
  };
}

function shuffled(values, seed) {
  const output = [...values];
  const random = seededRandom(seed);
  for (let index = output.length - 1; index > 0; index -= 1) {
    const swap = Math.floor(random() * (index + 1));
    [output[index], output[swap]] = [output[swap], output[index]];
  }
  return output;
}

async function captureTree(root) {
  const result = [];
  async function visit(directory, relative) {
    const entries = await readdir(directory, { withFileTypes: true });
    for (const entry of entries.sort((a, b) => a.name.localeCompare(b.name))) {
      const target = path.join(directory, entry.name);
      const childRelative = path.join(relative, entry.name).split(path.sep).join("/");
      if (entry.isDirectory()) await visit(target, childRelative);
      else if (entry.isFile()) result.push({
        path: childRelative,
        sha256: await digestFile(target),
        bytes: (await stat(target)).size,
      });
    }
  }
  await visit(root, "");
  return result;
}

async function loadInputs() {
  return Promise.all([
    readJson(CASES_PATH),
    readJson(RUBRIC_PATH),
    readJson(COVERAGE_PATH),
    readJson(FIXTURES_PATH),
  ]);
}

async function verifyFixture(fixture) {
  const result = {
    id: fixture.id,
    path: fixture.path,
    expectedSha256: fixture.sha256,
    status: "missing",
  };
  if (!(await exists(fixture.path))) return result;
  result.actualSha256 = await digestFile(fixture.path);
  result.status = result.actualSha256 === fixture.sha256 ? "ok" : "hash-mismatch";
  return result;
}

async function gitOutput(args) {
  const result = await runProcess("git", args, { cwd: REPO_ROOT, timeout: 30_000 });
  if (result.code !== 0) fail(result.stderr || "git command failed");
  return result.stdout;
}

async function repositoryIdentity() {
  return {
    repoRoot: REPO_ROOT,
    branch: (await gitOutput(["branch", "--show-current"])).trim(),
    head: (await gitOutput(["rev-parse", "HEAD"])).trim(),
    originMain: (await gitOutput(["rev-parse", "origin/main"])).trim(),
    node: process.version,
    platform: process.platform + "-" + process.arch,
    codex: await commandVersion("codex", ["--version"]),
    officekit: path.join(REPO_ROOT, "bin", "officekit.mjs"),
    dirtyMainPreserved: true,
  };
}

async function commandVersion(command, args) {
  const result = await runProcess(command, args, { cwd: REPO_ROOT, timeout: 30_000 });
  return result.code === 0 ? result.stdout.trim() : "unavailable";
}

function runProcess(command, args, options = {}) {
  const cwd = options.cwd || REPO_ROOT;
  const input = options.input;
  const timeout = Number(options.timeout || DEFAULT_TIMEOUT_MS);
  const stdoutFile = options.stdoutFile;
  const stderrFile = options.stderrFile;
  return new Promise((resolve) => {
    const startedAt = Date.now();
    const child = spawn(command, args, {
      cwd,
      env: options.env || process.env,
      stdio: ["pipe", "pipe", "pipe"],
      detached: Boolean(options.detached),
    });
    const stdoutParts = [];
    const stderrParts = [];
    let stdoutBytes = 0;
    let stderrBytes = 0;
    let timedOut = false;
    let outputStream;
    let errorStream;
    const openFiles = Promise.all([
      stdoutFile ? import("node:fs").then((module) => { outputStream = module.createWriteStream(stdoutFile); }) : Promise.resolve(),
      stderrFile ? import("node:fs").then((module) => { errorStream = module.createWriteStream(stderrFile); }) : Promise.resolve(),
    ]);
    openFiles.then(() => {
      child.stdout.on("data", (chunk) => {
        if (outputStream) outputStream.write(chunk);
        if (stdoutBytes < MAX_CAPTURE_BYTES) stdoutParts.push(chunk.subarray(0, Math.min(chunk.byteLength, MAX_CAPTURE_BYTES - stdoutBytes)));
        stdoutBytes += chunk.byteLength;
      });
      child.stderr.on("data", (chunk) => {
        if (errorStream) errorStream.write(chunk);
        if (stderrBytes < MAX_CAPTURE_BYTES) stderrParts.push(chunk.subarray(0, Math.min(chunk.byteLength, MAX_CAPTURE_BYTES - stderrBytes)));
        stderrBytes += chunk.byteLength;
      });
      child.stdin.end(input === undefined ? "" : input);
    });
    const timer = setTimeout(() => {
      timedOut = true;
      if (options.detached && child.pid) {
        try { process.kill(-child.pid, "SIGTERM"); } catch {}
        setTimeout(() => {
          try { process.kill(-child.pid, "SIGKILL"); } catch {}
        }, 5000).unref();
      } else {
        child.kill("SIGTERM");
        setTimeout(() => child.kill("SIGKILL"), 5000).unref();
      }
    }, timeout);
    child.on("error", (error) => {
      clearTimeout(timer);
      outputStream?.end();
      errorStream?.end();
      resolve({ command, args, cwd, code: null, signal: null, timedOut, error: error.message, stdout: "", stderr: error.message, durationMs: Date.now() - startedAt });
    });
    child.on("close", (code, signal) => {
      clearTimeout(timer);
      outputStream?.end();
      errorStream?.end();
      resolve({
        command, args, cwd, code, signal, timedOut,
        stdout: Buffer.concat(stdoutParts).toString("utf8"),
        stderr: Buffer.concat(stderrParts).toString("utf8"),
        stdoutBytes, stderrBytes, durationMs: Date.now() - startedAt,
      });
    });
  });
}

function armOrderForCase(seed, caseId, arms = ["shared-what-kind-how", "kimi-concise"]) {
  const caseSeed = seed + [...caseId].reduce((sum, character) => sum + character.charCodeAt(0), 0);
  return shuffled(arms, caseSeed >>> 0);
}

export async function validateExperiment(options = {}) {
  const [casesManifest, rubric, coverage, fixtures] = await loadInputs();
  const errors = [];
  const warnings = [];
  if (!["office-kit/presentation-skill-ablation-cases/v1", "office-kit/presentation-skill-ablation-cases/v2", "office-kit/presentation-skill-ablation-cases/v3"].includes(casesManifest.schema)) errors.push("unexpected cases schema");
  if (rubric.schema !== "office-kit/presentation-skill-ablation-rubric/v1") errors.push("unexpected rubric schema");
  if (coverage.schema !== "office-kit/presentation-skill-ablation-capability-coverage/v1") errors.push("unexpected coverage schema");
  if (fixtures.schema !== "office-kit/presentation-skill-ablation-fixtures/v1") errors.push("unexpected fixtures schema");
  if (casesManifest.model?.id !== "gpt-5.6-luna" || casesManifest.model?.reasoningEffort !== "max") errors.push("model must be gpt-5.6-luna/max");
  if (casesManifest.model?.timeoutMs !== 1200000) errors.push("frozen timeout must be 1200000ms");
  const expectedArmCount = casesManifest.schema.endsWith("/v3") ? 4 : casesManifest.schema.endsWith("/v2") ? 3 : 2;
  if (!Array.isArray(casesManifest.arms) || casesManifest.arms.length !== expectedArmCount) errors.push("expected " + expectedArmCount + " arms");
  if (casesManifest.schema.endsWith("/v2") && !casesManifest.arms.includes("current-production")) errors.push("v2 must include the current-production control arm");
  if (casesManifest.schema.endsWith("/v3") && !casesManifest.arms.includes("current-production")) errors.push("v3 must include the current-production control arm");
  const scoreKeys = new Set();
  if (!Array.isArray(rubric.dimensions)) errors.push("rubric dimensions must be an array");
  else {
    for (const dimension of rubric.dimensions) {
      if (typeof dimension.scoreKey !== "string" || !dimension.scoreKey) errors.push("rubric dimension " + dimension.id + " must declare scoreKey");
      else if (scoreKeys.has(dimension.scoreKey)) errors.push("duplicate rubric scoreKey " + dimension.scoreKey);
      else scoreKeys.add(dimension.scoreKey);
    }
  }
  const cases = casesManifest.cases || [];
  if (cases.length !== 12) errors.push("expected 12 cases, found " + cases.length);
  const scenarios = new Set(cases.map((item) => item.scenario));
  for (const scenario of ["analysis-decision", "management-report", "technical-engineering", "academic-research", "education-training", "brand-creative"]) {
    if (!scenarios.has(scenario)) errors.push("missing scenario " + scenario);
    const matching = cases.filter((item) => item.scenario === scenario);
    if (matching.filter((item) => item.lifecycle === "0-to-1").length !== 1) errors.push(scenario + " must have one 0-to-1 case");
    if (matching.filter((item) => item.lifecycle === "1-to-10").length !== 1) errors.push(scenario + " must have one 1-to-10 case");
  }
  const ids = new Set();
  for (const item of cases) {
    if (!item.id || ids.has(item.id)) errors.push("duplicate or missing case id " + (item.id || "<missing>"));
    ids.add(item.id);
    if (!item.brief || item.brief.length < 80) errors.push(item.id + ": brief is too short");
    if (!Array.isArray(item.requiredCapabilities) || item.requiredCapabilities.length < 3) errors.push(item.id + ": capability contract is too small");
    if (item.lifecycle === "1-to-10" && (!item.source?.fixtureId || item.source.targetPage < 1 || item.editSteps?.length !== 2)) errors.push(item.id + ": 1-to-10 source/edit contract is incomplete");
    if (item.lifecycle === "0-to-1" && item.source) errors.push(item.id + ": 0-to-1 case cannot have a source");
  }
  const declared = new Set(Object.keys(coverage.assignments || {}));
  for (const capability of REQUIRED_CAPABILITIES) {
    if (!declared.has(capability)) errors.push("capability ledger does not declare " + capability);
    if (!cases.some((item) => item.requiredCapabilities?.includes(capability))) errors.push("no case exercises " + capability);
  }
  const weights = Object.values(rubric.dimensions || {}).reduce((sum, item) => sum + Number(item.weight || 0), 0);
  if (weights !== 100) errors.push("rubric weights must total 100, found " + weights);
  const fixtureMap = new Map((fixtures.sources || []).map((item) => [item.id, item]));
  const fixtureResults = [];
  for (const item of cases.filter((candidate) => candidate.lifecycle === "1-to-10")) {
    if (!fixtureMap.has(item.source.fixtureId)) errors.push(item.id + ": fixture is absent");
  }
  if (options.checkFixtures !== false) {
    for (const fixture of fixtures.sources || []) {
      const result = await verifyFixture(fixture);
      fixtureResults.push(result);
      if (result.status === "missing") warnings.push("fixture missing: " + fixture.path);
      if (result.status === "hash-mismatch") errors.push("fixture hash mismatch: " + fixture.id);
    }
  }
  const commonTree = await captureTree(COMMON_ROOT);
  const armTrees = {};
  for (const arm of casesManifest.arms || []) {
    const armRoot = path.join(ARMS_ROOT, arm);
    if (!(await exists(path.join(armRoot, "SKILL.md")))) errors.push("missing arm " + arm);
    armTrees[arm] = await captureTree(armRoot);
  }
  for (const relative of REQUIRED_COMMON) if (!(await exists(path.join(COMMON_ROOT, relative)))) errors.push("missing common reference " + relative);
  const cleanRoomFiles = [...commonTree.map((item) => path.join(COMMON_ROOT, item.path)), ...(casesManifest.arms || []).map((arm) => path.join(ARMS_ROOT, arm, "SKILL.md"))];
  const cleanRoomText = (await Promise.all(cleanRoomFiles.map((file) => readFile(file, "utf8")))).join("\n").toLowerCase();
  for (const marker of FORBIDDEN_CLEAN_ROOM_MARKERS) if (cleanRoomText.includes(marker)) errors.push("clean-room overlay contains " + marker);
  const productionDiff = await gitOutput(["diff", "--name-only", "origin/main", "--", "src", "proto", "skills", "packages", "native"]).catch(() => "");
  if (productionDiff.trim()) errors.push("production paths changed in experiment branch:\n" + productionDiff.trim());
  if (errors.length) fail(errors.join("\n"));
  return {
    schema: "office-kit/presentation-skill-ablation-validation/v1",
    checkedAt: nowIso(),
    cases: cases.length,
    scenarios: [...scenarios].sort(),
    capabilities: [...REQUIRED_CAPABILITIES].sort(),
    common: { files: commonTree.length, sha256: digestText(JSON.stringify(commonTree)) },
    arms: Object.fromEntries(Object.entries(armTrees).map(([arm, tree]) => [arm, { files: tree.length, sha256: digestText(JSON.stringify(tree)) }])),
    fixtures: fixtureResults,
    warnings,
    ready: true,
  };
}

async function makeRunRoot(flags) {
  const supplied = flag(flags, "run-root");
  if (supplied) {
    const resolved = path.resolve(String(supplied));
    await ensureDir(resolved);
    return resolved;
  }
  const stamp = nowIso().replace(/[-:.TZ]/gu, "").slice(0, 14);
  const directory = path.join(EXPERIMENT_ROOT, "runs", "study-" + stamp + "-" + process.pid);
  await ensureDir(directory);
  return directory;
}

async function prepare(flags) {
  const validation = await validateExperiment({ checkFixtures: !boolFlag(flags, "skip-fixtures") });
  const [casesManifest, rubric, coverage, fixtures] = await loadInputs();
  const runRoot = await makeRunRoot(flags);
  const identity = await repositoryIdentity();
  const schedule = casesManifest.cases.flatMap((item) => armOrderForCase(casesManifest.seed || DEFAULT_SEED, item.id, casesManifest.arms).map((arm, index) => ({
    caseId: item.id,
    arm,
    order: index + 1,
  })));
  await writeJson(path.join(runRoot, "study.json"), {
    schema: "office-kit/presentation-skill-ablation-study/" + (casesManifest.arms.length === 4 ? "v3" : casesManifest.arms.length === 3 ? "v2" : "v1"),
    preparedAt: nowIso(),
    seed: casesManifest.seed || DEFAULT_SEED,
    model: casesManifest.model,
    arms: casesManifest.arms,
    identity,
    manifests: {
      casesSha256: await digestFile(CASES_PATH),
      rubricSha256: await digestFile(RUBRIC_PATH),
      coverageSha256: await digestFile(COVERAGE_PATH),
      fixturesSha256: await digestFile(FIXTURES_PATH),
      commonSha256: validation.common.sha256,
      armSha256: Object.fromEntries(Object.entries(validation.arms).map(([arm, value]) => [arm, value.sha256])),
    },
    schedule,
    fixtureResults: validation.fixtures,
    rawRuns: "outside-package-or-ignored",
  });
  await writeJson(path.join(runRoot, "frozen", "cases.v" + (casesManifest.arms.length === 4 ? "3" : casesManifest.arms.length === 3 ? "2" : "1") + ".json"), casesManifest);
  await writeJson(path.join(runRoot, "frozen", "rubric.v1.json"), rubric);
  await writeJson(path.join(runRoot, "frozen", "capability-coverage.v1.json"), coverage);
  await writeJson(path.join(runRoot, "frozen", "fixtures.v1.json"), fixtures);
  return { runRoot, validation, schedule };
}

function sourceForCase(item, fixtures) {
  if (!item.source) return null;
  const fixture = fixtures.sources.find((candidate) => candidate.id === item.source.fixtureId);
  if (!fixture) fail(item.id + ": fixture " + item.source.fixtureId + " is absent");
  return {
    ...item.source,
    fixtureId: fixture.id,
    path: fixture.path,
    sha256: fixture.sha256,
  };
}

function authorPrompt({ arm, item, skillPath, workspace, source, designReference }) {
  const lifecycle = item.lifecycle === "0-to-1" ? "0→1" : "1→10";
  const sourceInstruction = source
    ? "\nSource PPTX (read-only; never overwrite): " + source.path + "\nSource SHA-256: " + source.sha256 + "\nTarget page: " + source.targetPage + "\nUse officekit ppj import and perform exactly two serial edits: semantic first, then visual/delivery.\n"
    : "\nThere is no supplied page skeleton. Create exactly one dense representative page from the brief.\n";
  const referenceInstruction = designReference
    ? "\nOptional design reference to observe (do not copy its content): " + designReference + "\n"
    : "";
  return [
    "You are the author for a frozen presentation Skill experiment. Your arm is: " + arm + ". The lifecycle is " + lifecycle + ". Do not compare arms or mention this arm in the artifact.",
    "",
    "Read the arm instructions first: " + skillPath,
    "Read the case file at " + path.join(workspace, "input", "case.json"),
    "Use the OfficeKit repository at " + REPO_ROOT + "; invoke its public CLI as: node " + path.join(REPO_ROOT, "bin", "officekit.mjs") + " ppj ...",
    "Work only inside this trial workspace: " + workspace + ". Do not edit tracked repository files, do not use MJS/JSX as a presentation authoring surface, and do not run the full test suite.",
    sourceInstruction,
    referenceInstruction,
    "Follow the brief and acceptance contract exactly. Use only supplied facts; mark illustrative or assumed content. If a photo/icon is required, use the shared image route and record query, source, rights, hash, crop and alt text. Do not use host image generation.",
    "",
    "Required outputs (even if a check fails, leave a precise report):",
    "- outputs/deck.ppj",
    "- outputs/deck.pptx when the codec is available",
    "- outputs/previews/ (target page and adjacent page for 1→10; the page for 0→1)",
    "- outputs/review.json or the CLI review output",
    "- outputs/author-report.md explaining decisions, commands, failures and evidence type",
    "",
    "Run the narrow sequence check → build → render → repair visible defects → review. For 1→10 also re-import the output and explain stable IDs, opaque content, source binding and non-target preservation. Never declare structural evidence to be PowerPoint playback evidence. Stop after the requested case; do not create a benchmark harness.",
  ].join("\n");
}

async function prepareAuthorWorkspace(runRoot, arm, item, fixtures) {
  const workspace = path.join(runRoot, "authors", item.id, arm);
  await ensureDir(path.join(workspace, "input"));
  await ensureDir(path.join(workspace, "outputs"));
  await ensureDir(path.join(workspace, "evidence"));
  const source = sourceForCase(item, fixtures);
  const designReference = item.designReference ? path.join(REPO_ROOT, item.designReference) : null;
  const skillPath = path.join(ARMS_ROOT, arm, "SKILL.md");
  await writeJson(path.join(workspace, "input", "case.json"), { ...item, source, designReference });
  await writeJson(path.join(workspace, "input", "provenance.json"), {
    schema: "office-kit/presentation-skill-ablation-author-input/v1",
    caseId: item.id,
    arm,
    createdAt: nowIso(),
    skillPath,
    skillSha256: await digestFile(skillPath),
    source,
    designReference,
  });
  await writeFile(path.join(workspace, "input", "AUTHOR_PROMPT.md"), authorPrompt({
    arm, item, skillPath, workspace, source, designReference,
  }), "utf8");
  return { workspace, source, skillPath };
}

function eventStats(text) {
  const stats = {
    jsonlEvents: 0,
    fileReads: 0,
    toolCalls: 0,
    imageSearches: 0,
    retries: 0,
    checkPasses: 0,
    buildPasses: 0,
    renderPasses: 0,
    reviewPasses: 0,
    tokens: {},
  };
  for (const line of text.split(/\r?\n/).filter(Boolean)) {
    let event;
    try {
      event = JSON.parse(line);
    } catch {
      continue;
    }
    stats.jsonlEvents += 1;
    const serialized = JSON.stringify(event);
    const type = String(event.type || event.item?.type || "").toLowerCase();
    if (type.includes("file") && (type.includes("read") || type.includes("search"))) stats.fileReads += 1;
    if (type.includes("tool") || ["command_execution", "file_change", "mcp_tool_call", "web_search"].includes(event.item?.type)) stats.toolCalls += 1;
    if (/officekit[^\n]*image\s+search|image\s+search/iu.test(serialized)) stats.imageSearches += 1;
    if (/retry|retries/iu.test(type)) stats.retries += 1;
    if (/officekit[^\n]*(ppj\s+)?check/iu.test(serialized)) stats.checkPasses += 1;
    if (/officekit[^\n]*(ppj\s+)?build/iu.test(serialized)) stats.buildPasses += 1;
    if (/officekit[^\n]*(ppj\s+)?render/iu.test(serialized)) stats.renderPasses += 1;
    if (/officekit[^\n]*(ppj\s+)?review/iu.test(serialized)) stats.reviewPasses += 1;
    for (const [key, aliases] of Object.entries({
      inputTokens: ["input_tokens", "inputTokens"],
      outputTokens: ["output_tokens", "outputTokens"],
      reasoningTokens: ["reasoning_tokens", "reasoningTokens"],
    })) {
      for (const alias of aliases) {
        if (Number.isFinite(event.usage?.[alias])) stats.tokens[key] = (stats.tokens[key] || 0) + event.usage[alias];
      }
    }
    if (Number.isFinite(event.usage?.total_tokens)) stats.tokens.totalTokens = (stats.tokens.totalTokens || 0) + event.usage.total_tokens;
  }
  return stats;
}

async function findFirst(root, names) {
  for (const name of names) {
    const target = path.join(root, name);
    if (await exists(target)) return target;
  }
  return null;
}

async function runCli(commandArgs, cwd, evidenceDir, timeout) {
  const safeName = commandArgs.slice(1).join("-").replace(/[^a-z0-9_-]+/giu, "-").replace(/^-|-$/gu, "").slice(0, 80) || "command";
  const stdoutPath = path.join(evidenceDir, safeName + ".stdout");
  const stderrPath = path.join(evidenceDir, safeName + ".stderr");
  const result = await runProcess("node", [path.join(REPO_ROOT, "bin", "officekit.mjs"), ...commandArgs], {
    cwd, timeout, stdoutFile: stdoutPath, stderrFile: stderrPath,
  });
  return { ...result, stdoutPath, stderrPath };
}

function collectIds(value, output = []) {
  if (!value || typeof value !== "object") return output;
  if (Array.isArray(value)) {
    for (const child of value) collectIds(child, output);
    return output;
  }
  if (typeof value.id === "string") output.push(value.id);
  for (const child of Object.values(value)) collectIds(child, output);
  return output;
}

function collectStableIds(program) {
  const output = [];
  const visitElement = (element) => {
    if (!element || typeof element !== "object" || Array.isArray(element)) return;
    if (typeof element.id === "string") output.push(element.id);
    for (const key of ["elements", "children", "items"]) {
      if (Array.isArray(element[key])) for (const child of element[key]) visitElement(child);
    }
  };
  for (const page of program?.pages || []) {
    if (typeof page?.id === "string") output.push(page.id);
    for (const element of page?.elements || []) visitElement(element);
  }
  return output;
}

async function auditPpj({ workspace, item, source, timeout }) {
  const outputs = path.join(workspace, "outputs");
  const evidence = path.join(workspace, "evidence");
  await ensureDir(evidence);
  const ppj = await findFirst(outputs, ["deck.ppj", "output.ppj"]);
  const record = {
    schema: "office-kit/presentation-skill-ablation-artifact-evidence/v1",
    caseId: item.id,
    lifecycle: item.lifecycle,
    checkedAt: nowIso(),
    paths: { ppj, pptx: null },
    commands: [],
    gates: {},
    source: source ? { ...source, status: "unverified" } : null,
    stableIds: { status: "unverified" },
    metrics: {},
  };
  if (!ppj) {
    record.gates.schema = "missing";
    record.gates.overall = "failed";
    await writeJson(path.join(evidence, "artifact-evidence.json"), record);
    return record;
  }
  let ppjValue;
  try {
    ppjValue = JSON.parse(await readFile(ppj, "utf8"));
    record.metrics.ppjBytes = (await stat(ppj)).size;
    const ids = collectStableIds(ppjValue);
    record.stableIds = {
      status: new Set(ids).size === ids.length ? "ok" : "duplicate",
      count: ids.length,
    };
  } catch {
    record.gates.schema = "invalid-json";
  }
  const check = await runCli(["ppj", "check", ppj, "--json"], workspace, evidence, timeout);
  record.commands.push({ name: "check", code: check.code, durationMs: check.durationMs, stdoutPath: check.stdoutPath, stderrPath: check.stderrPath });
  record.gates.schema = check.code === 0 ? "passed" : (record.gates.schema || "failed");
  const runnerBuildPath = path.join(evidence, "runner-build.pptx");
  // The public CLI is deliberately fail-closed when an output already exists.
  // These paths are disposable audit products, so clear only those paths
  // before a repeatable re-audit instead of treating stale evidence as a gate.
  await rm(runnerBuildPath, { force: true });
  const build = await runCli(["ppj", "build", ppj, "-o", runnerBuildPath, "--json"], workspace, evidence, timeout);
  record.commands.push({ name: "build", code: build.code, durationMs: build.durationMs, stdoutPath: build.stdoutPath, stderrPath: build.stderrPath });
  record.gates.build = build.code === 0 ? "passed" : "failed";
  const finalPptx = await findFirst(outputs, ["deck.pptx", "output.pptx"]);
  if (finalPptx) {
    record.paths.pptx = finalPptx;
    record.metrics.pptxBytes = (await stat(finalPptx)).size;
    const pageCount = Number(ppjValue?.pages?.length || 1);
    const target = Number(source?.targetPage || 1);
    const first = item.lifecycle === "1-to-10" ? Math.max(1, target - 1) : 1;
    const last = item.lifecycle === "1-to-10" ? Math.min(pageCount, target + 1) : 1;
    const runnerRenderPath = path.join(evidence, "runner-rendered");
    await rm(runnerRenderPath, { recursive: true, force: true });
    const render = await runCli(["ppj", "render", ppj, "-o", runnerRenderPath, "--pages", first + "-" + last, "--json"], workspace, evidence, timeout);
    record.commands.push({ name: "render", code: render.code, durationMs: render.durationMs, stdoutPath: render.stdoutPath, stderrPath: render.stderrPath });
    record.gates.render = render.code === 0 ? "passed" : "failed";
    const review = await runCli(["ppj", "review", ppj, "--json"], workspace, evidence, timeout);
    record.commands.push({ name: "review", code: review.code, durationMs: review.durationMs, stdoutPath: review.stdoutPath, stderrPath: review.stderrPath });
    record.gates.review = review.code === 0 ? "passed" : "failed";
    if (item.lifecycle === "1-to-10") {
      const reimportPath = path.join(evidence, "reimport.ppj");
      const reimport = await runCli(["ppj", "import", finalPptx, "-o", reimportPath, "--json"], workspace, evidence, timeout);
      record.commands.push({ name: "reimport", code: reimport.code, durationMs: reimport.durationMs, stdoutPath: reimport.stdoutPath, stderrPath: reimport.stderrPath });
      record.gates.reimport = reimport.code === 0 ? "passed" : "failed";
      if (reimport.code === 0 && await exists(reimportPath)) {
        try {
          const imported = JSON.parse(await readFile(reimportPath, "utf8"));
          const originalIds = new Set(collectStableIds(ppjValue));
          const importedIds = new Set(collectStableIds(imported));
          const overlap = [...originalIds].filter((id) => importedIds.has(id)).length;
          record.stableIds.reimportOverlap = overlap + "/" + originalIds.size;
          record.stableIds.status = overlap === originalIds.size ? "ok" : "partial";
        } catch (error) {
          record.stableIds.reimportError = error.message;
        }
      }
      if (source) {
        record.source.actualSha256 = await digestFile(source.path).catch(() => null);
        record.source.status = record.source.actualSha256 === source.sha256 ? "source-hash-ok" : "source-hash-mismatch";
      }
    }
  } else {
    record.gates.render = "not-run-no-pptx";
    record.gates.review = "not-run-no-pptx";
  }
  const required = [record.gates.schema, record.gates.build, record.gates.render, record.gates.review];
  record.gates.overall = required.every((value) => value === "passed") &&
    record.stableIds.status === "ok" &&
    (!source || record.source.status === "source-hash-ok") ? "passed" : "failed";
  await writeJson(path.join(evidence, "artifact-evidence.json"), record);
  return record;
}

async function runOneAuthor({ runRoot, arm, item, fixtures, timeout }) {
  const prepared = await prepareAuthorWorkspace(runRoot, arm, item, fixtures);
  const prompt = await readFile(path.join(prepared.workspace, "input", "AUTHOR_PROMPT.md"), "utf8");
  const eventPath = path.join(prepared.workspace, "evidence", "codex-events.jsonl");
  const finalPath = path.join(prepared.workspace, "outputs", "codex-final.txt");
  const fixtureDir = prepared.source ? path.dirname(prepared.source.path) : REPO_ROOT;
  const args = [
    "exec", "--ephemeral", "--ignore-user-config", "--ignore-rules", "--json",
    "--sandbox", "workspace-write", "--skip-git-repo-check",
    "--add-dir", fixtureDir, "--model", "gpt-5.6-luna",
    "--config", "model_reasoning_effort=\"max\"",
    "--output-last-message", finalPath,
  ];
  const startedAt = nowIso();
  const result = await runProcess("codex", args, {
    cwd: prepared.workspace,
    input: prompt,
    timeout,
    detached: true,
    stdoutFile: eventPath,
    stderrFile: path.join(prepared.workspace, "evidence", "codex-stderr.log"),
  });
  const events = await readFile(eventPath, "utf8").catch(() => "");
  const artifact = await auditPpj({
    workspace: prepared.workspace,
    item,
    source: prepared.source,
    timeout,
  });
  const record = {
    schema: "office-kit/presentation-skill-ablation-author-run/v1",
    caseId: item.id,
    scenario: item.scenario,
    lifecycle: item.lifecycle,
    arm,
    startedAt,
    finishedAt: nowIso(),
    workspace: prepared.workspace,
    skillSha256: await digestFile(prepared.skillPath),
    codex: {
      model: "gpt-5.6-luna",
      reasoningEffort: "max",
      timeoutMs: timeout,
      code: result.code,
      signal: result.signal,
      timedOut: result.timedOut,
      durationMs: result.durationMs,
      finalPath,
      eventPath,
    },
    usage: eventStats(events),
    artifact,
  };
  await writeJson(path.join(prepared.workspace, "evidence", "author-run.json"), record);
  return record;
}

async function authorRecords(runRoot) {
  const result = [];
  async function visit(directory) {
    if (!(await exists(directory))) return;
    for (const entry of await readdir(directory, { withFileTypes: true })) {
      const target = path.join(directory, entry.name);
      if (entry.isDirectory()) await visit(target);
      else if (entry.isFile() && entry.name === "author-run.json") result.push(await readJson(target));
    }
  }
  await visit(path.join(runRoot, "authors"));
  return result;
}

async function runAuthors(flags) {
  const [casesManifest, , , fixtures] = await loadInputs();
  const runRoot = flag(flags, "run-root") ? path.resolve(String(flag(flags, "run-root"))) : (await prepare(flags)).runRoot;
  await ensureDir(runRoot);
  const requestedCase = flag(flags, "case");
  const requestedArm = flag(flags, "arm");
  const limit = flag(flags, "limit") ? positiveInteger(flag(flags, "limit"), "limit") : Infinity;
  const concurrency = flag(flags, "concurrency") ? positiveInteger(flag(flags, "concurrency"), "concurrency") : 1;
  const skipExisting = boolFlag(flags, "skip-existing");
  const targets = [];
  for (const item of casesManifest.cases) {
    if (requestedCase && item.id !== requestedCase) continue;
    for (const arm of armOrderForCase(casesManifest.seed || DEFAULT_SEED, item.id, casesManifest.arms)) {
      if (!requestedArm || requestedArm === arm) targets.push({ item, arm });
    }
  }
  if (!targets.length) fail("no author targets match filters");
  const selected = [];
  for (const target of targets.slice(0, limit)) {
    if (skipExisting && await exists(path.join(runRoot, "authors", target.item.id, target.arm, "evidence", "author-run.json"))) continue;
    selected.push(target);
  }
  if (!selected.length) return { runRoot, completed: 0, records: [] };
  const records = [];
  let cursor = 0;
  async function worker() {
    while (cursor < selected.length) {
      const target = selected[cursor++];
      process.stderr.write("author " + target.item.id + " / " + target.arm + "\n");
      const record = await runOneAuthor({ runRoot, arm: target.arm, item: target.item, fixtures, timeout: flag(flags, "timeout-ms", DEFAULT_TIMEOUT_MS) });
      records.push(record);
    }
  }
  await Promise.all(Array.from({ length: Math.min(concurrency, selected.length) }, () => worker()));
  const authorIndexPath = path.join(runRoot, "authors-index.json");
  const previousIndex = await exists(authorIndexPath) ? await readJson(authorIndexPath) : { records: [] };
  const recordPaths = new Set(previousIndex.records || []);
  for (const record of records) recordPaths.add(path.relative(runRoot, path.join(record.workspace, "evidence", "author-run.json")));
  await writeJson(authorIndexPath, {
    schema: "office-kit/presentation-skill-ablation-authors-index/v1",
    updatedAt: nowIso(),
    records: [...recordPaths].sort(),
  });
  return { runRoot, completed: records.length, records };
}

async function auditExisting(flags) {
  const [casesManifest, , , fixtures] = await loadInputs();
  const runRoot = path.resolve(String(flag(flags, "run-root", path.join(EXPERIMENT_ROOT, "runs"))));
  const caseId = flag(flags, "case");
  const arm = flag(flags, "arm");
  if (!caseId || !arm) fail("audit requires --case and --arm");
  const item = casesManifest.cases.find((candidate) => candidate.id === caseId);
  if (!item) fail("unknown case " + caseId);
  const workspace = path.join(runRoot, "authors", caseId, arm);
  const source = sourceForCase(item, fixtures);
  const artifact = await auditPpj({ workspace, item, source, timeout: flag(flags, "timeout-ms", DEFAULT_TIMEOUT_MS) });
  const authorRecordPath = path.join(workspace, "evidence", "author-run.json");
  if (await exists(authorRecordPath)) {
    const authorRecord = await readJson(authorRecordPath);
    authorRecord.artifact = artifact;
    authorRecord.artifactReauditedAt = nowIso();
    await writeJson(authorRecordPath, authorRecord);
  }
  console.log(JSON.stringify({ runRoot, caseId, arm, artifact }, null, 2));
  return artifact;
}

function judgeSchemaPath(slots = ["A", "B"]) {
  if (process.env.PRESENTATION_ABLATION_JUDGE_SCHEMA) {
    return path.join(EXPERIMENT_ROOT, process.env.PRESENTATION_ABLATION_JUDGE_SCHEMA);
  }
  const version = slots.length >= 4 ? "v3" : slots.length === 3 ? "v2" : "v1";
  return path.join(EXPERIMENT_ROOT, "judge-output-schema." + version + ".json");
}

function judgePrompt({ item, round, hardGate, slots = ["A", "B"] }) {
  const scoreFields = slots.map((slot) => slot + ":{communicationEvidence,hierarchyReadabilityOccupancy,designCraftSpecificity,functionalEditability,layerLayoutRender,displayMediumFit,completionPolish}").join(",");
  const gateFields = slots.map((slot) => slot + ':"passed|failed|unknown"').join(",");
  return [
    "You are a blind evaluator for a presentation quality study. Do not infer or mention which Skill produced any slot.",
    "Case: " + item.id + "; scenario=" + item.scenario + "; lifecycle=" + item.lifecycle,
    "Brief: " + item.brief,
    "",
    "Inspect attached rendered pages and, when present, local structural evidence. Slots " + slots.join(", ") + " are randomized labels and do not identify an authoring route.",
    "Score each dimension independently from 1 to 5. Anchor 1 = undeliverable or communication-breaking; 3 = usable but ordinary with a clear local issue; 5 = clear information, intentional design, reliable structure, close to production delivery.",
    "Weights: communication/evidence 20; hierarchy/readability/occupancy 15; design craft/creative specificity 20; functional/editability 15; layer/layout/render robustness 15; display/medium fit 10; completion/polish 5.",
    "Treat any structural hard-gate failure shown in evidence as a failure regardless of appearance. The redacted oracle summary is " + slots.map((slot) => slot + "=" + (hardGate?.[slot] || "unknown")).join(", ") + ". Do not claim playback unless evidence explicitly records a real host.",
    "",
    "This is blind round " + round + ". Return JSON only: {caseId, round, scores:{" + scoreFields + "},winner:\"" + slots.join("|") + "|tie|invalid\",confidence:1-5,hardGate:{" + gateFields + "},reason:\"brief evidence-based reason\"}.",
  ].join("\n");
}

function parseJudgeOutput(text, slots = ["A", "B"]) {
  const candidates = [text.trim(), ...(text.match(/\{[\s\S]*\}/gu) || [])];
  for (const candidate of candidates.reverse()) {
    try {
      const value = JSON.parse(candidate);
      if (slots.every((slot) => value?.scores?.[slot]) && value.winner) return value;
    } catch {
      // The final response may contain prose around the JSON. Keep it invalid
      // instead of silently inventing a score.
    }
  }
  return null;
}

async function prepareBlindSet(runRoot, item, records, round, seed, arms) {
  const byArm = new Map(records.filter((record) => record.caseId === item.id).map((record) => [record.arm, record]));
  if (!arms.every((arm) => byArm.has(arm))) return null;
  const slots = arms.map((_, index) => String.fromCharCode(65 + index));
  const order = shuffled(arms, seed + round * 1009 + item.id.length);
  const assignment = Object.fromEntries(slots.map((slot, index) => [slot, order[index]]));
  const destination = path.join(runRoot, "blind-review", "round-" + round, item.id);
  await ensureDir(destination);
  const files = {};
  for (const slot of slots) {
    const record = byArm.get(assignment[slot]);
    const source = path.join(record.workspace, "outputs", "previews");
    const target = path.join(destination, slot);
    await ensureDir(target);
    if (await exists(source)) await cp(source, target, { recursive: true, force: true });
    files[slot] = {
      directory: target,
      sourceArm: record.arm,
      evidence: record.artifact,
    };
  }
  await writeJson(path.join(destination, "pair.json"), {
    schema: "office-kit/presentation-skill-ablation-blind-set/v1",
    caseId: item.id,
    round,
    slots,
    assignment: Object.fromEntries(slots.map((slot) => [slot, "redacted"])),
    files: Object.fromEntries(slots.map((slot) => [slot, files[slot].directory])),
    createdAt: nowIso(),
  });
  // Keep the randomized label-to-arm mapping outside the judge workspace.
  // This lets analysis recover route-level deltas without exposing the arm
  // identity in the files presented to the blind evaluator.
  await writeJson(path.join(runRoot, "blind-review", "truth", "round-" + round, item.id + ".json"), {
    schema: "office-kit/presentation-skill-ablation-blind-truth/v1",
    caseId: item.id,
    round,
    slots,
    assignment,
    createdAt: nowIso(),
  });
  return { destination, assignment, files, slots, item };
}

async function findPreviews(directory) {
  if (!(await exists(directory))) return [];
  const names = (await readdir(directory)).sort();
  return names
    .filter((name) => /\.(png|jpe?g)$/iu.test(name))
    .map((name) => path.join(directory, name));
}

async function runOneJudge({ pair, round, timeout }) {
  const workspace = path.join(pair.destination, "judge");
  await ensureDir(workspace);
  const hardGate = Object.fromEntries(pair.slots.map((slot) => [slot, pair.files[slot].evidence?.gates?.overall === "passed" ? "passed" : "failed"]));
  const prompt = judgePrompt({ item: pair.item, round, hardGate, slots: pair.slots });
  const promptPath = path.join(workspace, "JUDGE_PROMPT.md");
  const eventPath = path.join(workspace, "events.jsonl");
  const finalPath = path.join(workspace, "final.txt");
  await writeFile(promptPath, prompt, "utf8");
  const args = [
    "exec", "--ephemeral", "--ignore-user-config", "--ignore-rules", "--json",
    "--sandbox", "read-only", "--skip-git-repo-check", "--model", "gpt-5.6-luna",
    "--config", "model_reasoning_effort=\"max\"", "--output-schema", judgeSchemaPath(pair.slots),
    "--output-last-message", finalPath,
  ];
  for (const slot of pair.slots) {
    for (const preview of await findPreviews(pair.files[slot].directory)) args.push("--image", preview);
  }
  const startedAt = nowIso();
  const result = await runProcess("codex", args, {
    cwd: workspace,
    input: prompt,
    timeout,
    detached: true,
    stdoutFile: eventPath,
    stderrFile: path.join(workspace, "stderr.log"),
  });
  const finalText = await readFile(finalPath, "utf8").catch(() => "");
  const parsed = parseJudgeOutput(finalText, pair.slots);
  if (parsed) parsed.hardGate = hardGate;
  const record = {
    schema: "office-kit/presentation-skill-ablation-judge-run/v1",
    caseId: pair.item.id,
    round,
    startedAt,
    finishedAt: nowIso(),
    assignment: Object.fromEntries(pair.slots.map((slot) => [slot, "redacted"])),
    codex: {
      model: "gpt-5.6-luna",
      reasoningEffort: "max",
      timeoutMs: timeout,
      code: result.code,
      durationMs: result.durationMs,
      eventPath,
      finalPath,
    },
    usage: eventStats(await readFile(eventPath, "utf8").catch(() => "")),
    score: parsed,
    status: parsed ? "valid" : "invalid",
  };
  await writeJson(path.join(pair.destination, "judge.json"), record);
  return record;
}

async function loadJudgeRecords(runRoot) {
  const result = [];
  async function visit(directory) {
    if (!(await exists(directory))) return;
    for (const entry of await readdir(directory, { withFileTypes: true })) {
      const target = path.join(directory, entry.name);
      if (entry.isDirectory()) await visit(target);
      else if (entry.isFile() && entry.name === "judge.json") result.push(await readJson(target));
    }
  }
  await visit(path.join(runRoot, "blind-review"));
  return result;
}

async function loadPairTruth(runRoot, caseId, round) {
  const file = path.join(runRoot, "blind-review", "truth", "round-" + round, caseId + ".json");
  return await exists(file) ? readJson(file) : null;
}

async function runJudges(flags) {
  const [casesManifest] = await loadInputs();
  const runRoot = path.resolve(String(flag(flags, "run-root", path.join(EXPERIMENT_ROOT, "runs"))));
  const authors = await authorRecords(runRoot);
  const roundStart = positiveInteger(flag(flags, "round", 1), "round");
  const rounds = boolFlag(flags, "both-rounds") ? [1, 2] : [roundStart];
  const requestedCase = flag(flags, "case");
  const limit = flag(flags, "limit") ? positiveInteger(flag(flags, "limit"), "limit") : Infinity;
  const concurrency = flag(flags, "concurrency") ? positiveInteger(flag(flags, "concurrency"), "concurrency") : 1;
  const skipExisting = boolFlag(flags, "skip-existing");
  const selected = [];
  for (const round of rounds) {
    let completed = 0;
    for (const item of casesManifest.cases) {
      if (requestedCase && requestedCase !== item.id) continue;
      const pair = await prepareBlindSet(runRoot, item, authors, round, casesManifest.seed || DEFAULT_SEED, casesManifest.arms);
      if (!pair) continue;
      if (skipExisting && await exists(path.join(pair.destination, "judge.json"))) continue;
      selected.push({ pair, round });
      completed += 1;
      if (completed >= limit) break;
    }
  }
  const results = [];
  let cursor = 0;
  async function worker() {
    while (cursor < selected.length) {
      const target = selected[cursor++];
      results.push(await runOneJudge({ pair: target.pair, round: target.round, timeout: flag(flags, "timeout-ms", DEFAULT_TIMEOUT_MS) }));
    }
  }
  await Promise.all(Array.from({ length: Math.min(concurrency, selected.length) }, () => worker()));
  await writeJson(path.join(runRoot, "judges-index.json"), {
    schema: "office-kit/presentation-skill-ablation-judges-index/v1",
    updatedAt: nowIso(),
    records: results.map((record) => path.join("blind-review", "round-" + record.round, record.caseId, "judge.json")),
  });
  return { runRoot, completed: results.length, records: results };
}

function weightedScore(scores, dimensions) {
  if (!scores) return null;
  let total = 0;
  for (const config of dimensions) {
    const value = Number(scores[config.scoreKey]);
    if (!Number.isFinite(value) || value < 1 || value > 5) return null;
    total += (value / 5) * Number(config.weight);
  }
  return Number(total.toFixed(3));
}

function mean(values) {
  return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null;
}

function median(values) {
  if (!values.length) return null;
  const sorted = [...values].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
}

function binomialCoefficient(n, k) {
  if (k < 0 || k > n) return 0;
  let result = 1;
  for (let index = 1; index <= Math.min(k, n - k); index += 1) result = result * (n - index + 1) / index;
  return result;
}

function signTest(values) {
  const nonzero = values.filter((value) => value !== 0);
  const n = nonzero.length;
  if (!n) return { n: 0, positive: 0, negative: 0, pTwoSided: 1 };
  const positive = nonzero.filter((value) => value > 0).length;
  const negative = n - positive;
  const smaller = Math.min(positive, negative);
  let tail = 0;
  for (let index = 0; index <= smaller; index += 1) tail += binomialCoefficient(n, index) / (2 ** n);
  return { n, positive, negative, pTwoSided: Math.min(1, 2 * tail) };
}

function exactPairedPermutation(values) {
  const clean = values.filter((value) => Number.isFinite(value));
  const n = clean.length;
  if (!n) return { n: 0, method: "exact-sign-flip", pTwoSided: 1 };
  if (n > 24) return { n, method: "exact-sign-flip-capped", pTwoSided: signTest(clean).pTwoSided };
  const observed = Math.abs(mean(clean));
  const combinations = 2 ** n;
  let atLeast = 0;
  for (let mask = 0; mask < combinations; mask += 1) {
    let total = 0;
    for (let index = 0; index < n; index += 1) total += (mask & (2 ** index)) === 0 ? clean[index] : -clean[index];
    if (Math.abs(total / n) + 1e-12 >= observed) atLeast += 1;
  }
  return { n, method: "exact-sign-flip", combinations, pTwoSided: atLeast / combinations };
}

function hardGatePairPassed(pair) {
  const statuses = Object.values(pair.hardGate || {});
  return statuses.length === 2 && statuses.every((status) => status === "passed");
}

function bootstrap(values, seed = DEFAULT_SEED, iterations = 2000) {
  if (!values.length) return { n: 0, low: null, high: null, median: null };
  const random = seededRandom(seed);
  const samples = [];
  for (let iteration = 0; iteration < iterations; iteration += 1) {
    let total = 0;
    for (let index = 0; index < values.length; index += 1) total += values[Math.floor(random() * values.length)];
    samples.push(total / values.length);
  }
  samples.sort((a, b) => a - b);
  return {
    n: values.length,
    low: samples[Math.floor(samples.length * 0.025)],
    high: samples[Math.floor(samples.length * 0.975)],
    median: median(values),
  };
}

function combinations(arms) {
  const result = [];
  for (let left = 0; left < arms.length; left += 1) {
    for (let right = left + 1; right < arms.length; right += 1) result.push([arms[left], arms[right]]);
  }
  return result;
}

function triPairwiseStats(scope, arms) {
  const output = {};
  for (const [left, right] of combinations(arms)) {
    const key = left + "__vs__" + right;
    const deltas = scope.flatMap((pair) => pair.rounds
      .filter((round) => round.eligible && Number.isFinite(round.scores[left]) && Number.isFinite(round.scores[right]))
      .map((round) => Number((round.scores[left] - round.scores[right]).toFixed(3))));
    output[key] = {
      left,
      right,
      deltas,
      meanDelta: mean(deltas),
      medianDelta: median(deltas),
      bootstrap95: bootstrap(deltas),
      signTest: signTest(deltas),
      pairedPermutation: exactPairedPermutation(deltas),
    };
  }
  return output;
}

async function analyzeTri(flags) {
  const [casesManifest, rubric] = await loadInputs();
  const arms = casesManifest.arms;
  const slots = arms.map((_, index) => String.fromCharCode(65 + index));
  const runRoot = path.resolve(String(flag(flags, "run-root", path.join(EXPERIMENT_ROOT, "runs"))));
  const authors = await authorRecords(runRoot);
  const judges = await loadJudgeRecords(runRoot);
  const pairs = [];
  for (const item of casesManifest.cases) {
    const authorByArm = new Map(authors.filter((record) => record.caseId === item.id).map((record) => [record.arm, record]));
    const hardGate = Object.fromEntries(arms.map((arm) => [arm, authorByArm.get(arm)?.artifact?.gates?.overall || "missing"]));
    const rounds = [];
    for (const record of judges.filter((candidate) => candidate.caseId === item.id && candidate.status === "valid")) {
      const truth = await loadPairTruth(runRoot, item.id, record.round);
      const scores = {};
      const judgeHardGate = {};
      for (const slot of slots) {
        const arm = truth?.assignment?.[slot];
        if (arm) {
          scores[arm] = weightedScore(record.score.scores?.[slot], rubric.dimensions);
          judgeHardGate[arm] = record.score.hardGate?.[slot] || "missing";
        }
      }
      rounds.push({
        round: record.round,
        scores,
        judgeHardGate,
        winnerArm: truth?.assignment && slots.includes(record.score.winner) ? truth.assignment[record.score.winner] : record.score.winner === "tie" ? "tie" : null,
        confidence: record.score.confidence,
        eligible: arms.every((arm) => hardGate[arm] === "passed" && judgeHardGate[arm] === "passed" && Number.isFinite(scores[arm])),
      });
    }
    pairs.push({
      caseId: item.id,
      scenario: item.scenario,
      lifecycle: item.lifecycle,
      image: Boolean(item.imageNeed && item.imageNeed !== "none"),
      hardGate,
      rounds,
      authors: Object.fromEntries(arms.filter((arm) => authorByArm.has(arm)).map((arm) => [arm, {
        durationMs: authorByArm.get(arm).codex.durationMs,
        usage: authorByArm.get(arm).usage,
        artifactStatus: authorByArm.get(arm).artifact?.gates?.overall || "missing",
      }])),
    });
  }
  const pairwise = triPairwiseStats(pairs, arms);
  // Only count winners from eligible rounds.  A judge can still return a
  // useful explanation for a failed artifact, but that appearance must not
  // become a route win in the quality summary.
  const routeWins = Object.fromEntries([...arms, "tie"].map((arm) => [arm, 0]));
  const allRouteWins = Object.fromEntries([...arms, "tie"].map((arm) => [arm, 0]));
  for (const pair of pairs) for (const round of pair.rounds) {
    if (round.winnerArm && allRouteWins[round.winnerArm] !== undefined) allRouteWins[round.winnerArm] += 1;
    if (round.eligible && round.winnerArm && routeWins[round.winnerArm] !== undefined) routeWins[round.winnerArm] += 1;
  }
  const byLifecycle = Object.fromEntries(["0-to-1", "1-to-10"].map((lifecycle) => [lifecycle, triPairwiseStats(pairs.filter((pair) => pair.lifecycle === lifecycle), arms)]));
  const scenarios = [...new Set(casesManifest.cases.map((item) => item.scenario))].sort();
  const byScenario = Object.fromEntries(scenarios.map((scenario) => [scenario, triPairwiseStats(pairs.filter((pair) => pair.scenario === scenario), arms)]));
  const byImage = Object.fromEntries([["image", true], ["non-image", false]].map(([name, image]) => [name, triPairwiseStats(pairs.filter((pair) => pair.image === image), arms)]));
  const efficiency = Object.fromEntries(arms.map((arm) => {
    const records = authors.filter((record) => record.arm === arm);
    const average = (selector) => mean(records.map(selector).filter((value) => Number.isFinite(value)));
    return [arm, {
      n: records.length,
      wallTimeMs: average((record) => record.codex?.durationMs),
      inputTokens: average((record) => record.usage?.tokens?.inputTokens),
      outputTokens: average((record) => record.usage?.tokens?.outputTokens),
      toolCalls: average((record) => record.usage?.toolCalls),
      imageSearches: average((record) => record.usage?.imageSearches),
    }];
  }));
  const routeLabel = arms.length === 3 ? "tri-route" : "multi-route";
  const report = {
    schema: "office-kit/presentation-skill-ablation-" + routeLabel + "-analysis/v1",
    generatedAt: nowIso(),
    runRoot: path.relative(REPO_ROOT, runRoot) || ".",
    arms,
    sample: {
      cases: casesManifest.cases.length,
      authorSets: pairs.filter((pair) => arms.every((arm) => Object.hasOwn(pair.authors, arm))).length,
      validHardGateSets: pairs.filter((pair) => arms.every((arm) => pair.hardGate[arm] === "passed")).length,
      visualHardGateRounds: pairs.flatMap((pair) => pair.rounds).filter((round) => arms.every((arm) => round.judgeHardGate?.[arm] === "passed")).length,
      eligibleRounds: pairs.flatMap((pair) => pair.rounds).filter((round) => round.eligible).length,
      judgeRecords: judges.length,
      judgedSets: pairs.filter((pair) => pair.rounds.length > 0).length,
      authorTimeouts: authors.filter((record) => record.codex?.timedOut).length,
      hardGateFailures: pairs.flatMap((pair) => arms
        .filter((arm) => pair.hardGate[arm] !== "passed")
        .map((arm) => ({ caseId: pair.caseId, arm, status: pair.hardGate[arm] }))),
      visualHardGateFailures: pairs.flatMap((pair) => pair.rounds.flatMap((round) => arms
        .filter((arm) => round.judgeHardGate?.[arm] !== "passed")
        .map((arm) => ({ caseId: pair.caseId, round: round.round, arm, status: round.judgeHardGate?.[arm] || "missing" })))),
    },
    routeWins,
    allRouteWins,
    pairwise,
    byLifecycle,
    byScenario,
    byImage,
    efficiency,
    pairs,
    limitations: [
      arms.length + " 方盲评只把作者侧和盲评侧硬门槛都通过的回合纳入 pairwise 质量分差",
      "Structural/render evidence is not PowerPoint playback evidence",
      "Human calibration records are pending unless supplied separately",
    ],
  };
  await writeJson(path.join(runRoot, "analysis." + routeLabel + ".v1.json"), report);
  const evidenceDir = path.join(EXPERIMENT_ROOT, "evidence");
  await ensureDir(evidenceDir);
  await writeJson(path.join(evidenceDir, "summary." + routeLabel + ".v1.json"), report);
  await writeFile(path.join(EXPERIMENT_ROOT, "report." + routeLabel + ".v1.md"), renderTriReport(report), "utf8");
  return report;
}

function renderTriReport(report) {
  const routeCount = report.arms.length;
  const lines = [
    "# Presentation Skill " + routeCount + " 路线质量实验：研究报告",
    "",
    "生成时间：" + report.generatedAt,
    "",
    "> 当前生产 Skill、Shared What/What-kind/How、Kimi-style concise 与新增混合路线共用同一 PPJ、素材、渲染和复核能力；本报告只比较入口路由。",
    "",
    "## 样本",
    "",
    "- 任务数：" + report.sample.cases + "；路线：" + report.arms.join("、") + "；全路线作者集合：" + report.sample.authorSets + "。",
    "- 所有路线的作者侧硬门槛均通过的任务：" + report.sample.validHardGateSets + "。",
    "- 盲评认为所有路线均通过视觉/内容硬门槛的回合：" + report.sample.visualHardGateRounds + "；实际纳入质量差分的回合：" + report.sample.eligibleRounds + "。",
    "- 有效盲评记录：" + report.sample.judgeRecords + "。",
    "",
    "## 路线胜负（仅合格配对）",
    "",
  ];
  for (const [arm, wins] of Object.entries(report.routeWins)) lines.push("- " + arm + "：" + wins + "。");
  lines.push("", "未过滤硬门槛的评审胜负仅作为诊断保留：" + Object.entries(report.allRouteWins).map(([arm, wins]) => arm + "=" + wins).join("，") + "。", "这些诊断胜负不进入质量结论。", "");
  lines.push("", "## 两两质量分差（左路线 − 右路线）", "", "| 比较 | n | 均值 | 中位数 | bootstrap 95% | p(sign) |", "| --- | ---: | ---: | ---: | --- | ---: |");
  for (const value of Object.values(report.pairwise)) lines.push("| " + value.left + " − " + value.right + " | " + value.deltas.length + " | " + (value.meanDelta ?? "pending") + " | " + (value.medianDelta ?? "pending") + " | " + (value.bootstrap95.low ?? "pending") + " … " + (value.bootstrap95.high ?? "pending") + " | " + value.signTest.pTwoSided + " |");
  lines.push("", "## 分层结果", "", "以下只在作者侧和盲评侧所有路线均通过硬门槛的回合中统计。", "", "| 分层 | 比较 | n | 均值 | 中位数 |", "| --- | --- | ---: | ---: | ---: |");
  for (const [lifecycle, values] of Object.entries(report.byLifecycle)) for (const value of Object.values(values)) lines.push("| lifecycle=" + lifecycle + " | " + value.left + " − " + value.right + " | " + value.deltas.length + " | " + (value.meanDelta ?? "pending") + " | " + (value.medianDelta ?? "pending") + " |");
  for (const [scenario, values] of Object.entries(report.byScenario)) for (const value of Object.values(values)) lines.push("| scenario=" + scenario + " | " + value.left + " − " + value.right + " | " + value.deltas.length + " | " + (value.meanDelta ?? "pending") + " | " + (value.medianDelta ?? "pending") + " |");
  for (const [image, values] of Object.entries(report.byImage)) for (const value of Object.values(values)) lines.push("| asset=" + image + " | " + value.left + " − " + value.right + " | " + value.deltas.length + " | " + (value.meanDelta ?? "pending") + " | " + (value.medianDelta ?? "pending") + " |");
  lines.push("", "## 效率", "", "| 路线 | 作者数 | 平均 wall time (ms) | 平均输入 token | 平均工具调用 |", "| --- | ---: | ---: | ---: | ---: |");
  for (const [arm, value] of Object.entries(report.efficiency)) lines.push("| " + arm + " | " + value.n + " | " + (value.wallTimeMs ?? "pending") + " | " + (value.inputTokens ?? "pending") + " | " + (value.toolCalls ?? "pending") + " |");
  lines.push("", "作者超时：" + report.sample.authorTimeouts + " / " + (report.sample.cases * report.arms.length) + "。", "作者侧硬门槛失败：" + report.sample.hardGateFailures.map((item) => item.caseId + "/" + item.arm + " (" + item.status + ")").join("；") + "。", "盲评侧视觉/内容硬门槛失败：" + report.sample.visualHardGateFailures.map((item) => item.caseId + "/r" + item.round + "/" + item.arm + " (" + item.status + ")").join("；") + "。", "");
  lines.push(
    "## 结论（探索性）",
    "",
    "- 这是一轮新增混合路线的探索性四方比较；质量差分同时要求作者侧和盲评侧硬门槛通过，不能用未通过门槛的外观分数抵消工程失败。",
    "- 路线胜负（合格回合）为：" + Object.entries(report.routeWins).map(([arm, wins]) => arm + "=" + wins).join("、") + "。",
    "- 每一对路线的均值、bootstrap 区间和配对检验见上表；样本仍是冻结的 12 个任务，不能外推到所有 PPT 场景。",
    "- 新混合路线的价值判断应同时看质量、硬门槛、编辑保真和成本；如果只在 0→1 获胜而 1→10 退化，不应直接替换生产路由。",
    "",
  );
  lines.push("", "## 限制", "");
  for (const item of report.limitations) lines.push("- " + item);
  return lines.join("\n") + "\n";
}

async function analyze(flags) {
  const [casesManifest, rubric] = await loadInputs();
  if (casesManifest.arms.length >= 3) return analyzeTri(flags);
  const runRoot = path.resolve(String(flag(flags, "run-root", path.join(EXPERIMENT_ROOT, "runs"))));
  const authors = await authorRecords(runRoot);
  const judges = await loadJudgeRecords(runRoot);
  const dimensions = rubric.dimensions;
  const pairs = [];
  for (const item of casesManifest.cases) {
    const authorByArm = new Map(authors.filter((record) => record.caseId === item.id).map((record) => [record.arm, record]));
    const judgeByRound = judges.filter((record) => record.caseId === item.id && record.status === "valid");
    const hardGate = Object.fromEntries([...authorByArm].map(([arm, record]) => [arm, record.artifact?.gates?.overall || "missing"]));
    const roundScores = [];
    for (const record of judgeByRound) {
      const truth = await loadPairTruth(runRoot, item.id, record.round);
      const A = weightedScore(record.score.scores.A, dimensions);
      const B = weightedScore(record.score.scores.B, dimensions);
      const sharedScore = truth?.assignment?.A === "shared-what-kind-how" ? A : truth?.assignment?.B === "shared-what-kind-how" ? B : null;
      const kimiScore = truth?.assignment?.A === "kimi-concise" ? A : truth?.assignment?.B === "kimi-concise" ? B : null;
      const winnerArm = truth?.assignment && ["A", "B"].includes(record.score.winner)
        ? truth.assignment[record.score.winner]
        : record.score.winner === "tie" ? "tie" : null;
      roundScores.push({
        round: record.round,
        A,
        B,
        shared: sharedScore,
        kimi: kimiScore,
        delta: sharedScore !== null && kimiScore !== null ? Number((sharedScore - kimiScore).toFixed(3)) : null,
        winner: record.score.winner,
        winnerArm,
        confidence: record.score.confidence,
        eligible: A !== null && B !== null && sharedScore !== null && kimiScore !== null &&
          Object.values(hardGate).length === 2 && Object.values(hardGate).every((status) => status === "passed"),
      });
    }
    pairs.push({
      caseId: item.id,
      scenario: item.scenario,
      lifecycle: item.lifecycle,
      image: Boolean(item.imageNeed && item.imageNeed !== "none"),
      hardGate,
      rounds: roundScores,
      meanDelta: mean(roundScores.filter((record) => record.delta !== null).map((record) => record.delta)),
      authors: Object.fromEntries([...authorByArm].map(([arm, record]) => [arm, {
        durationMs: record.codex.durationMs,
        usage: record.usage,
        artifactStatus: record.artifact?.gates?.overall || "missing",
      }])),
    });
  }
  const validDeltas = pairs.flatMap((pair) => pair.rounds.filter((round) => round.eligible).map((round) => round.delta));
  const winners = pairs.reduce((result, pair) => {
    for (const round of pair.rounds) {
      if (round.winner === "A") result.A += 1;
      else if (round.winner === "B") result.B += 1;
      else if (round.winner === "tie") result.tie += 1;
    }
    return result;
  }, { A: 0, B: 0, tie: 0 });
  const routeWinners = pairs.reduce((result, pair) => {
    for (const round of pair.rounds) {
      if (round.winnerArm === "shared-what-kind-how") result.shared += 1;
      else if (round.winnerArm === "kimi-concise") result.kimi += 1;
      else if (round.winnerArm === "tie") result.tie += 1;
    }
    return result;
  }, { shared: 0, kimi: 0, tie: 0 });
  const byLifecycle = Object.fromEntries(["0-to-1", "1-to-10"].map((lifecycle) => {
    const values = pairs.filter((pair) => pair.lifecycle === lifecycle).flatMap((pair) => pair.rounds.filter((round) => round.eligible).map((round) => round.delta));
    return [lifecycle, { n: values.length, meanDelta: mean(values), medianDelta: median(values), bootstrap95: bootstrap(values) }];
  }));
  const scenarios = [...new Set(casesManifest.cases.map((item) => item.scenario))].sort();
  const byScenario = Object.fromEntries(scenarios.map((scenario) => {
    const values = pairs.filter((pair) => pair.scenario === scenario).flatMap((pair) => pair.rounds.filter((round) => round.eligible).map((round) => round.delta));
    return [scenario, { n: values.length, meanDelta: mean(values), medianDelta: median(values) }];
  }));
  const byImage = Object.fromEntries([["image", true], ["non-image", false]].map(([name, image]) => {
    const values = pairs.filter((pair) => pair.image === image).flatMap((pair) => pair.rounds.filter((round) => round.eligible).map((round) => round.delta));
    return [name, { n: values.length, meanDelta: mean(values), medianDelta: median(values) }];
  }));
  const efficiency = Object.fromEntries(["shared-what-kind-how", "kimi-concise"].map((arm) => {
    const records = authors.filter((record) => record.arm === arm);
    const value = (selector) => mean(records.map(selector).filter((entry) => Number.isFinite(entry)));
    return [arm, {
      n: records.length,
      wallTimeMs: value((record) => record.codex?.durationMs),
      inputTokens: value((record) => record.usage?.tokens?.inputTokens),
      outputTokens: value((record) => record.usage?.tokens?.outputTokens),
      toolCalls: value((record) => record.usage?.toolCalls),
      imageSearches: value((record) => record.usage?.imageSearches),
      retries: value((record) => record.usage?.retries),
      renderReviewPasses: value((record) => (record.usage?.renderPasses || 0) + (record.usage?.reviewPasses || 0)),
    }];
  }));
  const agreementGroups = [];
  for (const pair of pairs) {
    const valid = pair.rounds.filter((round) => ["shared-what-kind-how", "kimi-concise", "tie"].includes(round.winnerArm));
    if (valid.length >= 2) agreementGroups.push(valid[0].winnerArm === valid[1].winnerArm);
  }
  const report = {
    schema: "office-kit/presentation-skill-ablation-analysis/v1",
    generatedAt: nowIso(),
    runRoot: path.relative(REPO_ROOT, runRoot) || ".",
    sample: {
      cases: casesManifest.cases.length,
      validHardGatePairs: pairs.filter(hardGatePairPassed).length,
      judgeRecords: judges.length,
      authorPairs: pairs.filter((pair) => Object.keys(pair.authors).length === 2).length,
      judgedPairs: pairs.filter((pair) => pair.rounds.length > 0).length,
    },
    winners,
    routeWinners,
    overall: {
      deltas: validDeltas,
      meanDelta: mean(validDeltas),
      medianDelta: median(validDeltas),
      bootstrap95: bootstrap(validDeltas),
      signTest: signTest(validDeltas),
      pairedPermutation: exactPairedPermutation(validDeltas),
    },
    byLifecycle,
    byScenario,
    byImage,
    efficiency,
    judgeAgreement: { n: agreementGroups.length, same: agreementGroups.filter(Boolean).length },
    pairs,
    limitations: [
      "冻结设计为 n=12 探索性配对样本；当前已完成作者配对 " + pairs.filter((pair) => Object.keys(pair.authors).length === 2).length + "/12，盲评记录 " + judges.length + "/24",
      "No default Skill or production route is selected",
      "Structural/render evidence is not PowerPoint playback evidence",
      "Human calibration records are pending unless supplied separately",
    ],
  };
  await writeJson(path.join(runRoot, "analysis.v1.json"), report);
  const evidenceDir = path.join(EXPERIMENT_ROOT, "evidence");
  await ensureDir(evidenceDir);
  await writeJson(path.join(evidenceDir, "summary.v1.json"), report);
  await writeFile(path.join(EXPERIMENT_ROOT, "report.v1.md"), renderReport(report), "utf8");
  return report;
}

function renderReport(report) {
  const lines = [
    "# Presentation Skill 双路线质量实验：研究报告",
    "",
    "生成时间：" + report.generatedAt,
    "",
    "> 本报告只解释冻结的 12 个配对任务，不改变生产默认路线，也不把结构渲染证据描述成真实 PowerPoint 播放通过。",
    "",
    "## 样本与硬门槛",
    "",
    "- 任务数：" + report.sample.cases + "；两套 Skill；每个场景各一项 0→1 和 1→10。",
    "- 双方都通过硬门槛的配对数：" + report.sample.validHardGatePairs + "。",
    "- 已解析的盲评记录：" + report.sample.judgeRecords + "。",
    "",
    "## 质量结果",
    "",
    "- 盲评标签胜负（仅审计轨迹）：A " + report.winners.A + "，B " + report.winners.B + "，平局 " + report.winners.tie + "。",
    "- 路线胜负（恢复随机映射后）：Shared " + report.routeWinners.shared + "，Kimi " + report.routeWinners.kimi + "，平局 " + report.routeWinners.tie + "。",
    "- 有效配对分差（Shared−Kimi）均值：" + (report.overall.meanDelta ?? "pending") + "；中位数：" + (report.overall.medianDelta ?? "pending") + "。",
    "- 精确符号检验：" + JSON.stringify(report.overall.signTest) + "。",
    "- 精确配对符号翻转检验：" + JSON.stringify(report.overall.pairedPermutation) + "。",
    "- bootstrap 95% 区间：" + JSON.stringify(report.overall.bootstrap95) + "。",
    "",
    "## 按生命周期",
    "",
    "| 生命周期 | 样本 | 均值分差 Shared−Kimi | 中位数分差 | bootstrap 95% |",
    "| --- | ---: | ---: | ---: | --- |",
  ];
  for (const [name, value] of Object.entries(report.byLifecycle)) {
    lines.push("| " + name + " | " + value.n + " | " + (value.meanDelta ?? "pending") + " | " + (value.medianDelta ?? "pending") + " | " + (value.bootstrap95.low ?? "pending") + " … " + (value.bootstrap95.high ?? "pending") + " |");
  }
  lines.push("", "## 按场景", "", "| 场景 | 样本 | 均值分差 Shared−Kimi | 中位数分差 |", "| --- | ---: | ---: | ---: |");
  for (const [name, value] of Object.entries(report.byScenario)) lines.push("| " + name + " | " + value.n + " | " + (value.meanDelta ?? "pending") + " | " + (value.medianDelta ?? "pending") + " |");
  lines.push("", "## 按图片任务", "", "| 分层 | 样本 | 均值分差 Shared−Kimi | 中位数分差 |", "| --- | ---: | ---: | ---: |");
  for (const [name, value] of Object.entries(report.byImage)) lines.push("| " + name + " | " + value.n + " | " + (value.meanDelta ?? "pending") + " | " + (value.medianDelta ?? "pending") + " |");
  lines.push("", "## 效率（不并入质量分）", "", "| 路线 | 作者数 | 平均 wall time (ms) | 平均输入 token | 平均工具调用 | 平均图片搜索 |", "| --- | ---: | ---: | ---: | ---: | ---: |");
  for (const [name, value] of Object.entries(report.efficiency)) lines.push("| " + name + " | " + value.n + " | " + (value.wallTimeMs ?? "pending") + " | " + (value.inputTokens ?? "pending") + " | " + (value.toolCalls ?? "pending") + " | " + (value.imageSearches ?? "pending") + " |");
  lines.push("", "## 限制", "");
  for (const item of report.limitations) lines.push("- " + item);
  lines.push("", "原始运行记录位于本次 runRoot；只将本汇总和冻结研究方法纳入版本控制。", "");
  return lines.join("\n") + "\n";
}

async function createHumanCalibrationTemplate(flags) {
  const runRoot = path.resolve(String(flag(flags, "run-root", path.join(EXPERIMENT_ROOT, "runs"))));
  const authors = await authorRecords(runRoot);
  const [casesManifest] = await loadInputs();
  const slots = casesManifest.arms.map((_, index) => String.fromCharCode(65 + index));
  const candidates = casesManifest.cases.map((item) => item.id)
    .filter((id) => casesManifest.arms.every((arm) => authors.some((record) => record.caseId === id && record.arm === arm)))
    .slice(0, 4);
  const output = {
    schema: "office-kit/presentation-skill-ablation-human-calibration/" + (casesManifest.arms.length === 3 ? "v2" : "v1"),
    createdAt: nowIso(),
    status: candidates.length === 4 ? "pending-human-input" : "insufficient-pairs",
    instruction: "由人类评审填写；不要透露 A/B/C 对应的路线（双路线历史运行仍可只填 A/B）。人类结果不用于调参。",
    pairs: candidates.map((caseId, index) => ({
      caseId, pair: index + 1,
      scores: Object.fromEntries(slots.map((slot) => [slot, null])),
      winner: null, confidence: null, reason: null,
    })),
  };
  await writeJson(path.join(runRoot, "blind-review", "human-calibration.v1.json"), output);
  return output;
}

export async function main(argv = process.argv.slice(2)) {
  const parsed = parseArgs(argv);
  const command = parsed.command;
  const flags = parsed.flags;
  if (command === "validate" || command === "smoke") {
    const result = await validateExperiment({ checkFixtures: !boolFlag(flags, "skip-fixtures") });
    if (command === "smoke") {
      const [casesManifest] = await loadInputs();
      const commonText = (await Promise.all((await captureTree(COMMON_ROOT)).map((entry) => readFile(path.join(COMMON_ROOT, entry.path), "utf8")))).join("\n");
      const routeText = await Promise.all(casesManifest.arms.map(async (arm) => {
        const root = path.join(ARMS_ROOT, arm);
        const tree = await captureTree(root);
        return commonText + "\n" + (await Promise.all(tree.map((entry) => readFile(path.join(root, entry.path), "utf8")))).join("\n");
      }));
      for (const text of routeText) {
        if (!/check\s*→\s*build\s*→\s*render/iu.test(text) || !/review/iu.test(text)) fail("route smoke: check/build/render/review route is not reachable");
        if (!/occlusion/iu.test(text) || !/(?:source[- ]bound|source binding)/iu.test(text) || !/image/iu.test(text)) fail("route smoke: shared invariants are not reachable");
      }
      result.smoke = { arms: casesManifest.arms, mandatoryRoute: "check→build→render→review", sharedInvariants: true };
    }
    console.log(JSON.stringify(result, null, 2));
    return result;
  }
  if (command === "prepare") {
    const result = await prepare(flags);
    console.log(JSON.stringify(result, null, 2));
    return result;
  }
  if (command === "run-authors") {
    const result = await runAuthors(flags);
    console.log(JSON.stringify({ runRoot: result.runRoot, completed: result.completed }, null, 2));
    return result;
  }
  if (command === "audit") {
    return auditExisting(flags);
  }
  if (command === "run-judges") {
    const result = await runJudges(flags);
    console.log(JSON.stringify({ runRoot: result.runRoot, completed: result.completed }, null, 2));
    return result;
  }
  if (command === "human-template") {
    const result = await createHumanCalibrationTemplate(flags);
    console.log(JSON.stringify(result, null, 2));
    return result;
  }
  if (command === "analyze") {
    const result = await analyze(flags);
    console.log(JSON.stringify({ runRoot: result.runRoot, validHardGatePairs: result.sample.validHardGatePairs }, null, 2));
    return result;
  }
  console.log([
    "Usage:",
    "  node evals/presentation-skill-ablation/run.mjs validate [--skip-fixtures]",
    "  node evals/presentation-skill-ablation/run.mjs smoke",
    "  node evals/presentation-skill-ablation/run.mjs prepare [--run-root <dir>]",
    "  node evals/presentation-skill-ablation/run.mjs run-authors [--run-root <dir>] [--case <id>] [--arm <id>] [--limit N] [--concurrency N] [--skip-existing]",
    "  node evals/presentation-skill-ablation/run.mjs audit --run-root <dir> --case <id> --arm <id>",
    "  node evals/presentation-skill-ablation/run.mjs run-judges --run-root <dir> [--both-rounds] [--limit N] [--concurrency N] [--skip-existing]",
    "  node evals/presentation-skill-ablation/run.mjs human-template --run-root <dir>",
    "  node evals/presentation-skill-ablation/run.mjs analyze --run-root <dir>",
  ].join("\n"));
}

if (import.meta.url === "file://" + process.argv[1]) {
  main().catch((error) => {
    console.error(error.stack || error.message || String(error));
    process.exitCode = 1;
  });
}
