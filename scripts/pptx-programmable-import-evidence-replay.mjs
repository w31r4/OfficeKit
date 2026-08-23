#!/usr/bin/env node

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import {
  compareRenderedPages,
  evaluatePackageOracle,
  readIntentDefinitions,
  renderPresentationPages,
  sha256,
} from "./pptx-programmable-import-oracle.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const defaultDefinitions = path.join(repoRoot, "evals/pptx-programmable-import/intent-matrix.v1.json");

export async function replayMatrixEvidence({ definitionsPath, runRoot, inputEvidencePath, outputPath }) {
  const [definitions, inputEvidenceBytes] = await Promise.all([
    readIntentDefinitions(definitionsPath),
    readFile(inputEvidencePath),
  ]);
  const inputEvidence = JSON.parse(inputEvidenceBytes.toString("utf8"));
  if (inputEvidence?.schema !== "office-kit/pptx-programmable-import-matrix/v1") throw new Error("Invalid matrix evidence schema");
  if (inputEvidence.baseline !== definitions.baseline) throw new Error("Matrix evidence baseline disagrees with definitions");
  if (inputEvidence.acceptance?.requiredRuns !== 90 || inputEvidence.acceptance?.requiredIntents !== 30) throw new Error("Replay requires the complete 30 x 3 matrix");
  if (inputEvidence.environment?.render !== true) throw new Error("Replay requires the real render-backed matrix");

  const renderCache = path.join(runRoot, "render-cache");
  const sources = [];
  for (const source of definitions.sources) {
    const originalSource = inputEvidence.sources.find(({ id }) => id === source.id);
    if (!originalSource) throw new Error(`${source.id}: source is missing from input evidence`);
    const intents = [];
    for (const intent of source.intents) {
      const originalIntent = originalSource.intents.find(({ id }) => id === intent.id);
      if (!originalIntent) throw new Error(`${source.id}/${intent.id}: intent is missing from input evidence`);
      const runs = [];
      for (let repetition = 1; repetition <= 3; repetition += 1) {
        const label = `${source.id}/${intent.id}/${repetition}`;
        const originalRun = originalIntent.runs.find((run) => run.repetition === repetition);
        if (!originalRun) throw new Error(`${label}: original run is missing`);
        const runDir = path.join(runRoot, "runs", source.id, intent.id, String(repetition));
        const inputPath = path.join(runDir, "source.pptx");
        const outputPathForRun = path.join(runDir, "output.pptx");
        const [sourceBytes, outputBytes, worker] = await Promise.all([
          readFile(inputPath),
          readFile(outputPathForRun),
          readFile(path.join(runDir, "worker.json"), "utf8").then(JSON.parse),
        ]);
        const sourceSha256After = sha256(sourceBytes);
        const outputSha256 = sha256(outputBytes);
        if (sourceSha256After !== source.sha256) throw new Error(`${label}: source hash changed`);
        if (worker.sourceSha256 !== source.sha256 || worker.outputSha256 !== outputSha256 || worker.secondImport !== true) {
          throw new Error(`${label}: worker receipt is inconsistent`);
        }
        const packageOracle = await evaluatePackageOracle({ sourceBytes, outputBytes, source, intent });
        const [sourceRender, outputRender] = await Promise.all([
          renderPresentationPages(inputPath, renderCache, source.sha256),
          renderPresentationPages(outputPathForRun, renderCache, outputSha256),
        ]);
        let pixelOracle;
        try {
          pixelOracle = compareRenderedPages(sourceRender, outputRender, intent.targetPage);
        } catch (error) {
          pixelOracle = { passed: false, reason: errorMessage(error) };
        }
        const status = pixelOracle.passed ? "passed" : "failed";
        const reason = pixelOracle.passed ? undefined : pixelOracle.reason;
        if (originalRun.status !== status || (status === "failed" && originalRun.reason !== reason)) {
          throw new Error(`${label}: replay changed the original outcome (${originalRun.status} -> ${status})`);
        }
        if (originalRun.outputSha256 && originalRun.outputSha256 !== outputSha256) throw new Error(`${label}: output hash changed since the original evaluation`);
        runs.push({
          repetition,
          status,
          ...(reason ? { reason } : {}),
          sourceSha256After,
          outputSha256,
          worker,
          packageOracle,
          pixelOracle,
        });
      }
      const passing = runs.filter(({ status }) => status === "passed");
      const outputHashes = new Set(passing.map(({ outputSha256 }) => outputSha256));
      const oracleHashes = new Set(passing.map(({ packageOracle, pixelOracle }) => sha256(Buffer.from(JSON.stringify({
        packageOracle,
        pixelOracle: pixelOracle.passed ? { ...pixelOracle, outputCacheHit: undefined } : pixelOracle,
      })))));
      const deterministic = passing.length === 3 && outputHashes.size === 1 && oracleHashes.size === 1;
      if (deterministic !== originalIntent.deterministic) throw new Error(`${source.id}/${intent.id}: replay changed deterministic status`);
      intents.push({ ...originalIntent, passedRuns: passing.length, deterministic, runs });
    }
    sources.push({ ...originalSource, intents });
  }

  const allIntents = sources.flatMap(({ intents }) => intents);
  const replayed = {
    ...inputEvidence,
    replay: {
      mode: "independent-evaluator-only",
      originalEvidenceSha256: sha256(inputEvidenceBytes),
      editsRerun: false,
      outcomesChanged: false,
      partialChecksRetained: true,
    },
    sources,
    acceptance: {
      requiredIntents: allIntents.length,
      requiredRuns: allIntents.length * 3,
      passedRuns: allIntents.reduce((count, intent) => count + intent.passedRuns, 0),
      deterministicIntents: allIntents.filter(({ deterministic }) => deterministic).length,
      status: allIntents.every(({ deterministic }) => deterministic) ? "passed" : "failed",
    },
  };
  if (JSON.stringify(replayed.acceptance) !== JSON.stringify(inputEvidence.acceptance)) throw new Error("Replay changed the acceptance summary");
  await writeFile(outputPath, `${JSON.stringify(replayed, null, 2)}\n`, { flag: "wx" });
  return replayed;
}

function errorMessage(error) {
  return error instanceof Error ? error.message : String(error);
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

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const replayed = await replayMatrixEvidence({
    definitionsPath: path.resolve(args.definitions || defaultDefinitions),
    runRoot: path.resolve(required(args, "run-root")),
    inputEvidencePath: path.resolve(args.input || path.join(required(args, "run-root"), "evidence.json")),
    outputPath: path.resolve(required(args, "output")),
  });
  process.stdout.write(`${JSON.stringify({ output: path.resolve(args.output), acceptance: replayed.acceptance, replay: replayed.replay }, null, 2)}\n`);
}

if (import.meta.url === pathToFileURL(process.argv[1] || "").href) {
  await main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
