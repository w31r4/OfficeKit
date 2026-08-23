#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { constants } from "node:fs";
import { chmod, copyFile, mkdir, readFile, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  compareRenderedPages,
  evaluatePackageOracle,
  readIntentDefinitions,
  renderPresentationPages,
  sha256,
} from "./pptx-programmable-import-oracle.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const defaultDefinitions = path.join(repoRoot, "evals/pptx-programmable-import/intent-matrix.v1.json");

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const definitionsPath = path.resolve(args.definitions || defaultDefinitions);
  const definitions = await readIntentDefinitions(definitionsPath);
  const assetsDir = path.resolve(args["assets-dir"] || process.env[definitions.assetsEnvironment] || definitions.defaultAssetsDirectory);
  const runRoot = path.resolve(required(args, "run-root"));
  await requireAbsent(runRoot, "run root");
  await mkdir(runRoot, { recursive: true });
  const packageRoot = path.resolve(args["package-root"] || repoRoot);
  const packageMetadata = JSON.parse(await readFile(path.join(packageRoot, "package.json"), "utf8"));
  if (packageMetadata.name !== "office-kit") throw new Error(`--package-root is not office-kit: ${packageRoot}`);
  const officekitBin = path.join(packageRoot, packageMetadata.bin?.officekit || "bin/officekit.mjs");
  await stat(officekitBin);
  const render = args["no-render"] !== true;
  const repetitions = args.repetitions ? positiveInteger(args.repetitions, "repetitions") : definitions.repetitionsPerIntent;
  const selectedSources = args.source ? definitions.sources.filter(({ id }) => id === args.source) : definitions.sources;
  if (!selectedSources.length) throw new Error(`Unknown --source ${args.source}`);
  const renderCache = path.join(runRoot, "render-cache");
  await mkdir(renderCache);
  const sources = [];
  for (const source of selectedSources) {
    const sourcePath = path.join(assetsDir, source.fileName);
    const sourceBytes = await readFile(sourcePath);
    if (sha256(sourceBytes) !== source.sha256) throw new Error(`${source.id}: immutable source hash mismatch`);
    const selectedIntents = args.intent ? source.intents.filter(({ id }) => id === args.intent) : source.intents;
    if (!selectedIntents.length) throw new Error(`Unknown --intent ${source.id}/${args.intent}`);
    const sourceRender = render ? await renderPresentationPages(sourcePath, renderCache, source.sha256) : null;
    const intentResults = [];
    for (const intent of selectedIntents) {
      const runs = [];
      for (let repetition = 1; repetition <= repetitions; repetition += 1) {
        const runDir = path.join(runRoot, "runs", source.id, intent.id, String(repetition));
        await mkdir(runDir, { recursive: true });
        const inputPath = path.join(runDir, "source.pptx");
        const outputPath = path.join(runDir, "output.pptx");
        const receiptPath = path.join(runDir, "worker.json");
        await copyFile(sourcePath, inputPath, constants.COPYFILE_EXCL);
        await chmod(inputPath, 0o444);
        const workerResult = runWorker({ officekitBin, definitionsPath, source, intent, inputPath, outputPath, receiptPath, cwd: runDir });
        await writeFile(path.join(runDir, "worker.stdout.txt"), workerResult.stdout, { flag: "wx" });
        await writeFile(path.join(runDir, "worker.stderr.txt"), workerResult.stderr, { flag: "wx" });
        let record;
        try {
          if (workerResult.status !== 0) throw new Error(`public worker exited ${workerResult.status}: ${workerResult.stderr.trim() || workerResult.stdout.trim()}`);
          const [outputBytes, receipt] = await Promise.all([readFile(outputPath), readFile(receiptPath, "utf8").then(JSON.parse)]);
          const sourceAfter = sha256(await readFile(inputPath));
          if (sourceAfter !== source.sha256) throw new Error(`source copy changed: ${sourceAfter}`);
          const packageOracle = await evaluatePackageOracle({ sourceBytes, outputBytes, source, intent });
          const outputRender = render ? await renderPresentationPages(outputPath, renderCache, sha256(outputBytes)) : null;
          const pixelOracle = render ? compareRenderedPages(sourceRender, outputRender, intent.targetPage) : { passed: false, skipped: true, reason: "--no-render" };
          record = { repetition, status: "passed", sourceSha256After: sourceAfter, outputSha256: sha256(outputBytes), worker: receipt, packageOracle, pixelOracle };
        } catch (error) {
          record = { repetition, status: "failed", reason: error instanceof Error ? error.message : String(error) };
        }
        runs.push(record);
      }
      const passing = runs.filter(({ status }) => status === "passed");
      const outputHashes = new Set(passing.map(({ outputSha256 }) => outputSha256));
      const oracleHashes = new Set(passing.map(({ packageOracle, pixelOracle }) => sha256(Buffer.from(JSON.stringify({ packageOracle, pixelOracle: pixelOracle?.passed ? { ...pixelOracle, outputCacheHit: undefined } : pixelOracle })) )));
      const deterministic = passing.length === repetitions && outputHashes.size === 1 && oracleHashes.size === 1;
      intentResults.push({ id: intent.id, targetPage: intent.targetPage, requiredRuns: repetitions, completedRuns: runs.length, passedRuns: passing.length, deterministic, runs });
    }
    sources.push({ id: source.id, sourceSha256: source.sha256, intents: intentResults });
  }
  const allIntents = sources.flatMap((source) => source.intents);
  const evidence = {
    schema: "office-kit/pptx-programmable-import-matrix/v1",
    baseline: definitions.baseline,
    definitionsSha256: sha256(await readFile(definitionsPath)),
    package: {
      name: packageMetadata.name,
      version: packageMetadata.version,
      installKind: args["install-kind"] || (packageRoot === repoRoot ? "repository" : "packed-clean-install"),
      tarballSha256: args["tarball-sha256"] || null,
    },
    environment: { platform: process.platform, arch: process.arch, node: process.version, render },
    repetitionsPerIntent: repetitions,
    sources,
    acceptance: {
      requiredIntents: allIntents.length,
      requiredRuns: allIntents.length * repetitions,
      passedRuns: allIntents.reduce((count, intent) => count + intent.passedRuns, 0),
      deterministicIntents: allIntents.filter(({ deterministic }) => deterministic).length,
      status: allIntents.length > 0 && allIntents.every(({ deterministic }) => deterministic) ? "passed" : "failed",
    },
  };
  await writeFile(path.join(runRoot, "evidence.json"), `${JSON.stringify(evidence, null, 2)}\n`, { flag: "wx" });
  process.stdout.write(`${JSON.stringify({ evidence: path.join(runRoot, "evidence.json"), acceptance: evidence.acceptance }, null, 2)}\n`);
  if (evidence.acceptance.status !== "passed") process.exitCode = 1;
}

function runWorker({ officekitBin, definitionsPath, source, intent, inputPath, outputPath, receiptPath, cwd }) {
  const result = spawnSync(process.execPath, [
    officekitBin,
    "run",
    path.join(repoRoot, "scripts/pptx-programmable-import-worker.mjs"),
    "--definitions", definitionsPath,
    "--source-id", source.id,
    "--intent-id", intent.id,
    "--input", inputPath,
    "--output", outputPath,
    "--receipt", receiptPath,
  ], { cwd, encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
  return { status: result.status, stdout: result.stdout || "", stderr: result.stderr || "" };
}

function parseArgs(argv) {
  const result = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) throw new Error(`Unexpected argument ${token}`);
    const name = token.slice(2);
    if (name === "no-render") result[name] = true;
    else if (argv[index + 1] && !argv[index + 1].startsWith("--")) result[name] = argv[++index];
    else throw new Error(`Missing value for ${token}`);
  }
  return result;
}

function required(args, name) {
  if (!args[name]) throw new Error(`Missing --${name}`);
  return args[name];
}

function positiveInteger(value, name) {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 1 || parsed > 10) throw new Error(`${name} must be an integer from 1 through 10`);
  return parsed;
}

async function requireAbsent(target, label) {
  try {
    await stat(target);
  } catch (error) {
    if (error.code === "ENOENT") return;
    throw error;
  }
  throw new Error(`${label} already exists; outputs are create-only: ${target}`);
}

await main().catch((error) => {
  process.stderr.write(`${error?.stack || error}\n`);
  process.exitCode = 2;
});
