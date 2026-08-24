#!/usr/bin/env node

import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const MANIFEST = path.resolve(import.meta.dirname, "../evals/presentation-authoring-compiler/pilot.v1.json");

export function scorePilot(manifest, runs, judgments = []) {
  const byArm = Object.fromEntries(["A", "B", "C"].map((arm) => [arm, runs.filter((run) => run.arm === arm)]));
  const hard = Object.fromEntries(["A", "B", "C"].map((arm) => [arm, rate(byArm[arm], (run) => run.status === "passed")]));
  const continuation = rate(byArm.C, (run) => run.checks?.task?.passed === true);
  const blindOverA = blindWinRate(judgments, "C", "A");
  const blindOverB = blindWinRate(judgments, "C", "B");
  const costs = {
    A: medianCost(byArm.A),
    B: medianCost(byArm.B),
    C: medianCost(byArm.C),
  };
  const ratio = costs.A?.combined && costs.C?.combined ? costs.C.combined / costs.A.combined : null;
  const thresholds = {
    hardPassRate: threshold(hard.C, manifest.thresholds.hardPassRate, hard.C !== null),
    hardPassDeltaFromA: threshold(hard.C === null || hard.A === null ? null : hard.C - hard.A, manifest.thresholds.hardPassDeltaFromA, hard.C !== null && hard.A !== null),
    blindWinRateOverA: threshold(blindOverA, manifest.thresholds.blindWinRateOverA, blindOverA !== null),
    blindWinRateOverB: threshold(blindOverB, manifest.thresholds.blindWinRateOverB, blindOverB !== null),
    selectedContinuationSuccess: threshold(continuation, manifest.thresholds.selectedContinuationSuccess, continuation !== null),
    medianTimeAndTokensRatioToA: threshold(ratio, manifest.thresholds.medianTimeAndTokensRatioToA, ratio !== null),
  };
  const complete = Object.values(thresholds).every((entry) => entry.status === "passed");
  const expectedRuns = manifest.design.totalRuns;
  return {
    schema: "office-kit/presentation-authoring-pilot-results/v1",
    manifestSchema: manifest.schema,
    expectedRuns,
    observedRuns: runs.length,
    arms: Object.fromEntries(["A", "B", "C"].map((arm) => [arm, { runs: byArm[arm].length, hardPassRate: hard[arm], medianCost: costs[arm] }])),
    blind: { judgments: judgments.length, overA: blindOverA, overB: blindOverB },
    continuationSuccess: continuation,
    thresholds,
    rollout: {
      status: complete ? "switch-C" : "keep-A",
      shippedDefault: complete ? "C" : "A",
      experimental: complete ? null : "C",
      reason: complete
        ? "Every predeclared threshold passed on the frozen pilot."
        : runs.length < expectedRuns
          ? "Pilot is incomplete; no default switch is allowed."
          : "At least one predeclared threshold is missing or failed; keep A shipped.",
    },
  };
}

function rate(records, predicate) {
  if (!records.length) return null;
  return records.filter(predicate).length / records.length;
}

function medianCost(records) {
  const costs = records.map((run) => {
    const time = Number(run.elapsedMs);
    const tokens = Number(run.tokenUsage?.totalTokens);
    if (!Number.isFinite(time) || time <= 0 || !Number.isFinite(tokens) || tokens <= 0) return null;
    return { time, tokens, combined: time * tokens };
  }).filter(Boolean);
  if (!costs.length) return null;
  return {
    timeMs: median(costs.map((cost) => cost.time)),
    totalTokens: median(costs.map((cost) => cost.tokens)),
    combined: median(costs.map((cost) => cost.combined)),
  };
}

function median(values) {
  const sorted = [...values].sort((left, right) => left - right);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
}

function blindWinRate(judgments, candidate, baseline) {
  const relevant = judgments.filter((judgment) => {
    const left = [judgment.leftArm, judgment.rightArm];
    return left.includes(candidate) && left.includes(baseline) && candidate !== baseline;
  });
  if (!relevant.length) return null;
  let wins = 0;
  for (const judgment of relevant) {
    const candidateIsLeft = judgment.leftArm === candidate;
    if (judgment.winner === (candidateIsLeft ? "left" : "right")) wins += 1;
    if (judgment.winner === "tie") wins += 0.5;
  }
  return wins / relevant.length;
}

function threshold(actual, requirement, available) {
  if (!available || actual === null) return { status: "insufficient-evidence", actual, requirement };
  const passed = requirement.operator === ">="
    ? actual >= requirement.value
    : requirement.operator === "<="
      ? actual <= requirement.value
      : actual > requirement.value;
  return { status: passed ? "passed" : "failed", actual, requirement };
}

async function readJudgments(file) {
  if (!file) return [];
  const text = await readFile(file, "utf8");
  return text.split(/\r?\n/u).filter(Boolean).map((line) => JSON.parse(line));
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const manifest = JSON.parse(await readFile(MANIFEST, "utf8"));
  const runsPath = path.resolve(required(args, "runs"));
  const runsPayload = JSON.parse(await readFile(runsPath, "utf8"));
  const judgments = await readJudgments(args.judgments ? path.resolve(args.judgments) : null);
  const result = scorePilot(manifest, runsPayload.runs || [], judgments);
  const output = path.resolve(args.output || path.join(path.dirname(runsPath), "results.v1.json"));
  await writeFile(output, `${JSON.stringify(result, null, 2)}\n`);
  process.stdout.write(`${JSON.stringify({ output, rollout: result.rollout, thresholds: result.thresholds }, null, 2)}\n`);
  if (result.rollout.status !== "switch-C") process.exitCode = 1;
}

function parseArgs(argv) {
  const result = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) throw new Error(`Unexpected argument ${token}`);
    const name = token.slice(2);
    if (!argv[index + 1] || argv[index + 1].startsWith("--")) throw new Error(`Missing value for ${token}`);
    result[name] = argv[++index];
  }
  return result;
}

function required(args, name) {
  assert.ok(args[name], `Missing --${name}`);
  return args[name];
}

if (import.meta.url === pathToFileURL(process.argv[1] || "").href) {
  await main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
