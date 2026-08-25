#!/usr/bin/env node

import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdir, readFile, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import {
  loadPilotManifest,
  packCandidate,
  runPilotTrial,
} from "./presentation-authoring-pilot.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const expansionPath = path.join(repoRoot, "evals/presentation-authoring-compiler/expansion.v1.json");
const DEFAULT_TIMEOUT_MS = 30 * 60 * 1000;

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const expansion = JSON.parse(await readFile(expansionPath, "utf8"));
  assert.equal(expansion.schema, "office-kit/presentation-authoring-expansion/v1");
  assert.equal(expansion.route, "C");
  assert.equal(expansion.tasks.length, expansion.taskCount);
  const pilot = await loadPilotManifest();
  const runRoot = path.resolve(required(args, "run-root"));
  await requireAbsent(runRoot);
  await mkdir(runRoot, { recursive: true });
  const packageInfo = await packCandidate(runRoot);
  const concurrency = positiveInteger(args.concurrency || "1", "concurrency", 8);
  const timeoutMs = args["timeout-ms"]
    ? positiveInteger(args["timeout-ms"], "timeout-ms", 24 * 60 * 60 * 1000)
    : DEFAULT_TIMEOUT_MS;
  const codexBin = args.codex || process.env.OFFICEKIT_CODEX_BIN || "codex";
  const records = new Array(expansion.tasks.length);
  let nextIndex = 0;

  async function worker() {
    while (true) {
      const index = nextIndex++;
      if (index >= expansion.tasks.length) return;
      const task = expansion.tasks[index];
      try {
        records[index] = await runPilotTrial({
          manifest: pilot,
          packageInfo,
          runRoot,
          task,
          arm: "C",
          trial: 1,
          armOrder: ["C"],
          timeoutMs,
          codexBin,
        });
      } catch (error) {
        records[index] = {
          schema: "office-kit/presentation-authoring-pilot-run/v1",
          runId: `${task.id}/C/1`,
          taskId: task.id,
          scenario: task.scenario,
          arm: "C",
          trial: 1,
          freshContext: true,
          packedCleanInstall: true,
          elapsedMs: null,
          retryCount: 0,
          attempts: 1,
          tokenUsage: { observed: false, inputTokens: null, outputTokens: null, totalTokens: null },
          status: "failed",
          failures: [`runner-${sanitize(errorMessage(error), runRoot)}`],
          checks: {},
        };
      }
    }
  }

  await Promise.all(Array.from({ length: Math.min(concurrency, expansion.tasks.length) }, () => worker()));
  const compactRuns = records.map(compactRun);
  const result = {
    schema: "office-kit/presentation-authoring-expansion-runs/v1",
    expansionManifestSha256: sha256(await readFile(expansionPath)),
    route: "C",
    sourcePolicy: "brief-only",
    package: {
      name: packageInfo.name,
      version: packageInfo.version,
      tarballSha256: packageInfo.tarballSha256,
      packedBytes: packageInfo.packedBytes,
      unpackedBytes: packageInfo.unpackedBytes,
      totalFiles: packageInfo.totalFiles,
    },
    environment: {
      platform: process.platform,
      arch: process.arch,
      node: process.version,
    },
    runs: compactRuns,
    acceptance: {
      expectedRuns: expansion.tasks.length,
      completedRuns: compactRuns.length,
      passedRuns: compactRuns.filter((run) => run.status === "passed").length,
      status: compactRuns.length === expansion.tasks.length && compactRuns.every((run) => run.status === "passed") ? "passed" : "failed",
    },
  };
  const summaryPath = args.summary ? path.resolve(args.summary) : path.join(runRoot, "expansion-runs.v1.json");
  await writeFile(summaryPath, `${JSON.stringify(result, null, 2)}\n`, { flag: "wx" });
  process.stdout.write(`${JSON.stringify({ summary: summaryPath, acceptance: result.acceptance }, null, 2)}\n`);
  if (result.acceptance.status !== "passed") process.exitCode = 1;
}

function compactRun(run) {
  const policy = run.checks?.policy;
  const task = run.checks?.task;
  const output = run.checks?.output;
  return {
    schema: run.schema,
    runId: run.runId,
    taskId: run.taskId,
    scenario: run.scenario,
    arm: run.arm,
    trial: run.trial,
    freshContext: run.freshContext,
    packedCleanInstall: run.packedCleanInstall,
    elapsedMs: run.elapsedMs,
    retryCount: run.retryCount,
    attempts: run.attempts,
    tokenUsage: run.tokenUsage,
    status: run.status,
    failures: run.failures,
    checks: {
      policy: policy ? { passed: policy.passed, commandCount: policy.commandCount, authoredFileCount: policy.authoredFileCount, findings: policy.findings } : null,
      task: task ? { passed: task.passed, taskId: task.taskId, commits: task.commits, publications: task.publications, plan: task.plan } : null,
      output: output ? { passed: output.passed, slides: output.slides, reviewVerdict: output.reviewVerdict, visualReview: output.visualReview, designWarnings: output.designWarnings } : null,
    },
  };
}

function parseArgs(argv) {
  const result = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) throw new Error(`Unexpected argument ${token}`);
    const name = token.slice(2);
    if (argv[index + 1] && !argv[index + 1].startsWith("--")) result[name] = argv[++index];
    else throw new Error(`Missing value for ${token}`);
  }
  return result;
}

function required(args, name) {
  if (!args[name]) throw new Error(`Missing --${name}`);
  return args[name];
}

function positiveInteger(value, name, maximum) {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 1 || parsed > maximum) throw new Error(`${name} must be an integer from 1 through ${maximum}`);
  return parsed;
}

async function requireAbsent(target) {
  try { await stat(target); }
  catch (error) { if (error.code === "ENOENT") return; throw error; }
  throw new Error(`Evidence root already exists: ${target}`);
}

function sanitize(value, runRoot) {
  return String(value).replaceAll(runRoot, "<run-root>").replaceAll(repoRoot, "<repo>");
}

function errorMessage(error) {
  return error instanceof Error ? error.message : String(error);
}

function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

if (import.meta.url === pathToFileURL(process.argv[1] || "").href) {
  await main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
