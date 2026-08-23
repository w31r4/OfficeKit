#!/usr/bin/env node

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { sha256 } from "./pptx-programmable-import-oracle.mjs";

export function buildBaselineEvidence({ matrix, codex, matrixBytes, codexBytes, harnessHead }) {
  if (matrix?.schema !== "office-kit/pptx-programmable-import-matrix/v1") throw new Error("Invalid matrix evidence schema");
  if (codex?.schema !== "office-kit/pptx-codex-continuation-evidence/v1") throw new Error("Invalid Codex evidence schema");
  if (!/^[0-9a-f]{40}$/u.test(harnessHead)) throw new Error("Harness HEAD must be a full lowercase Git SHA");
  if (matrix.baseline !== codex.baseline) throw new Error("Matrix and Codex evidence use different product baselines");
  if (matrix.acceptance?.requiredIntents !== 30 || matrix.acceptance?.requiredRuns !== 90 || matrix.repetitionsPerIntent !== 3) {
    throw new Error("Matrix evidence is not the complete 30 x 3 run");
  }
  if (codex.acceptance?.requiredTasks !== 3 || codex.acceptance?.trialsPerTask !== 3 || codex.acceptance?.requiredTrials !== 9) {
    throw new Error("Codex evidence is not the complete 3 x 3 run");
  }
  if (matrix.acceptance.passedRuns > matrix.acceptance.requiredRuns || codex.acceptance.passedTrials > codex.acceptance.requiredTrials) {
    throw new Error("Evidence pass counts exceed required counts");
  }
  if (matrix.package?.name !== "office-kit" || matrix.package?.installKind !== "packed-clean-install") {
    throw new Error("Matrix did not use a packed office-kit clean install");
  }
  if (codex.package?.name !== "office-kit" || codex.package?.cleanInstallPerTrial !== true) {
    throw new Error("Codex trials did not use clean installs");
  }
  if (!matrix.package.tarballSha256 || matrix.package.tarballSha256 !== codex.package.tarballSha256) {
    throw new Error("Matrix and Codex trials did not use the same deterministic tarball");
  }
  validateMatrixShape(matrix);
  validateCodexShape(codex, matrix.package.tarballSha256);
  const matrixIntents = matrix.sources.flatMap((source) => source.intents.map((intent) => ({ sourceId: source.id, ...intent })));
  const nonDeterministic = matrixIntents.filter(({ deterministic }) => !deterministic).map(({ sourceId, id, passedRuns, runs }) => ({
    sourceId,
    intentId: id,
    passedRuns,
    failedRuns: runs.filter(({ status }) => status !== "passed").map(({ repetition, reason }) => ({ repetition, reason })),
    outputSha256s: [...new Set(runs.filter(({ status }) => status === "passed").map(({ outputSha256 }) => outputSha256))],
  }));
  const failedCodexTrials = codex.trials.filter(({ status }) => status !== "passed").map(({ taskId, sourceId, repetition, failures, checks }) => ({
    taskId,
    sourceId,
    repetition,
    failures,
    codexStatus: checks.codex?.status ?? null,
    agentFinal: checks.codex?.agentFinal ?? null,
    agentFinalTruncated: checks.codex?.agentFinalTruncated ?? false,
    policyPassed: checks.policy?.passed ?? false,
    outputPassed: checks.output?.passed ?? false,
  }));
  const matrixStatus = matrix.acceptance.status;
  const codexStatus = codex.acceptance.status;
  return {
    schema: "office-kit/pptx-programmable-import-baseline/v1",
    productBaseline: matrix.baseline,
    acceptanceHarnessHead: harnessHead,
    package: {
      name: matrix.package.name,
      version: matrix.package.version,
      tarballSha256: matrix.package.tarballSha256,
      packedCleanInstall: true,
      codexCleanInstallPerTrial: true,
    },
    evidenceFiles: {
      matrix: { path: "baseline/matrix.v1.json", sha256: sha256(matrixBytes) },
      codex: { path: "baseline/codex.v1.json", sha256: sha256(codexBytes) },
    },
    definitions: {
      matrixSha256: matrix.definitionsSha256,
      continuationSha256: codex.definitions.continuationSha256,
      intentSha256: codex.definitions.intentSha256,
    },
    matrix: {
      requiredIntents: matrix.acceptance.requiredIntents,
      requiredRuns: matrix.acceptance.requiredRuns,
      passedRuns: matrix.acceptance.passedRuns,
      deterministicIntents: matrix.acceptance.deterministicIntents,
      status: matrixStatus,
      nonDeterministic,
    },
    codex: {
      requiredTasks: codex.acceptance.requiredTasks,
      requiredTrials: codex.acceptance.requiredTrials,
      completedTrials: codex.acceptance.completedTrials,
      passedTrials: codex.acceptance.passedTrials,
      status: codexStatus,
      failedTrials: failedCodexTrials,
    },
    environment: {
      matrix: matrix.environment,
      codex: codex.environment,
    },
    acceptance: {
      status: matrixStatus === "passed" && codexStatus === "passed" ? "passed" : "failed",
      failuresPreserved: true,
      oracleWeakened: false,
      productModifiedByAcceptance: false,
    },
  };
}

function validateMatrixShape(matrix) {
  if (!Array.isArray(matrix.sources) || matrix.sources.length !== 3 || new Set(matrix.sources.map(({ id }) => id)).size !== 3) {
    throw new Error("Matrix evidence must contain three distinct sources");
  }
  const intents = [];
  for (const source of matrix.sources) {
    if (!Array.isArray(source.intents) || source.intents.length !== 10 || new Set(source.intents.map(({ id }) => id)).size !== 10) {
      throw new Error(`${source.id}: matrix evidence must contain ten distinct intents`);
    }
    for (const intent of source.intents) {
      if (!Array.isArray(intent.runs) || intent.runs.length !== 3) throw new Error(`${source.id}/${intent.id}: expected three runs`);
      if (intent.requiredRuns !== 3 || intent.completedRuns !== 3) throw new Error(`${source.id}/${intent.id}: incomplete run counters`);
      if (intent.runs.some(({ repetition }, index) => repetition !== index + 1)) throw new Error(`${source.id}/${intent.id}: repetition sequence is invalid`);
      const passedRuns = intent.runs.filter(({ status }) => status === "passed").length;
      if (intent.passedRuns !== passedRuns) throw new Error(`${source.id}/${intent.id}: passed-run counter is inconsistent`);
      const deterministic = passedRuns === 3 && new Set(intent.runs.map(({ outputSha256 }) => outputSha256)).size === 1;
      if (intent.deterministic && !deterministic) throw new Error(`${source.id}/${intent.id}: deterministic claim disagrees with output hashes`);
      intents.push(intent);
    }
  }
  const passedRuns = intents.reduce((count, intent) => count + intent.runs.filter(({ status }) => status === "passed").length, 0);
  const deterministicIntents = intents.filter(({ deterministic }) => deterministic).length;
  if (matrix.acceptance.requiredIntents !== intents.length || matrix.acceptance.requiredRuns !== intents.length * 3) {
    throw new Error("Matrix acceptance totals are inconsistent");
  }
  if (matrix.acceptance.passedRuns !== passedRuns || matrix.acceptance.deterministicIntents !== deterministicIntents) {
    throw new Error("Matrix acceptance result counters are inconsistent");
  }
  const expectedStatus = deterministicIntents === intents.length ? "passed" : "failed";
  if (matrix.acceptance.status !== expectedStatus) throw new Error("Matrix acceptance status is inconsistent");
}

function validateCodexShape(codex, tarballSha256) {
  if (!Array.isArray(codex.trials) || codex.trials.length !== 9) throw new Error("Codex evidence must contain nine trials");
  const grouped = new Map();
  for (const trial of codex.trials) {
    if (!grouped.has(trial.taskId)) grouped.set(trial.taskId, []);
    grouped.get(trial.taskId).push(trial);
    if (trial.freshCodexContext !== true) throw new Error(`${trial.taskId}/${trial.repetition}: Codex context is not fresh`);
    if (trial.packedCleanInstall?.passed !== true || trial.packedCleanInstall.package !== "office-kit" || trial.packedCleanInstall.tarballSha256 !== tarballSha256) {
      throw new Error(`${trial.taskId}/${trial.repetition}: packed clean-install evidence is invalid`);
    }
    if (!new Set(["passed", "failed"]).has(trial.status)) throw new Error(`${trial.taskId}/${trial.repetition}: invalid status`);
    if (trial.status === "failed" && (!Array.isArray(trial.failures) || trial.failures.length === 0)) {
      throw new Error(`${trial.taskId}/${trial.repetition}: failed trial has no preserved reason`);
    }
  }
  if (grouped.size !== 3) throw new Error("Codex evidence must contain three distinct tasks");
  for (const [taskId, trials] of grouped) {
    const repetitions = trials.map(({ repetition }) => repetition).sort((left, right) => left - right);
    if (trials.length !== 3 || repetitions.join(",") !== "1,2,3") throw new Error(`${taskId}: expected repetitions 1, 2, and 3`);
  }
  const passedTrials = codex.trials.filter(({ status }) => status === "passed").length;
  if (codex.acceptance.completedTrials !== codex.trials.length || codex.acceptance.passedTrials !== passedTrials) {
    throw new Error("Codex acceptance result counters are inconsistent");
  }
  const expectedStatus = passedTrials === codex.trials.length ? "passed" : "failed";
  if (codex.acceptance.status !== expectedStatus) throw new Error("Codex acceptance status is inconsistent");
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const matrixPath = path.resolve(required(args, "matrix"));
  const codexPath = path.resolve(required(args, "codex"));
  const outputPath = path.resolve(required(args, "output"));
  const [matrixBytes, codexBytes] = await Promise.all([readFile(matrixPath), readFile(codexPath)]);
  const baseline = buildBaselineEvidence({
    matrix: JSON.parse(matrixBytes.toString("utf8")),
    codex: JSON.parse(codexBytes.toString("utf8")),
    matrixBytes,
    codexBytes,
    harnessHead: required(args, "harness-head"),
  });
  await writeFile(outputPath, `${JSON.stringify(baseline, null, 2)}\n`, { flag: "wx" });
  process.stdout.write(`${JSON.stringify({ output: outputPath, acceptance: baseline.acceptance, matrix: baseline.matrix.status, codex: baseline.codex.status }, null, 2)}\n`);
}

function parseArgs(argv) {
  const result = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--") || !argv[index + 1] || argv[index + 1].startsWith("--")) throw new Error(`Expected --name value, received ${token}`);
    result[token.slice(2)] = argv[++index];
  }
  return result;
}

function required(args, name) {
  if (!args[name]) throw new Error(`Missing --${name}`);
  return args[name];
}

if (import.meta.url === pathToFileURL(process.argv[1] || "").href) {
  await main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
