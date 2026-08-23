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

export function buildCandidateEvidence({ matrix, codex, companion, matrixBytes, codexBytes, companionBytes, candidateHead }) {
  if (!/^[0-9a-f]{40}$/u.test(candidateHead)) throw new Error("Candidate HEAD must be a full lowercase Git SHA");
  if (matrix?.schema !== "office-kit/pptx-programmable-import-matrix/v1") throw new Error("Invalid matrix evidence schema");
  if (codex?.schema !== "office-kit/pptx-codex-continuation-evidence/v1") throw new Error("Invalid Codex evidence schema");
  if (companion?.schema !== "office-kit/pptx-source-derived-companion-evidence/v1") throw new Error("Invalid source-derived companion evidence schema");
  if (matrix.baseline !== codex.baseline || matrix.baseline !== companion.productBaseline) {
    throw new Error("Candidate evidence components use different product baselines");
  }
  if (matrix.acceptance?.requiredIntents !== 30 || matrix.acceptance?.requiredRuns !== 90 || matrix.repetitionsPerIntent !== 3) {
    throw new Error("Matrix evidence is not the complete 30 x 3 run");
  }
  if (codex.acceptance?.requiredTasks !== 3 || codex.acceptance?.trialsPerTask !== 3 || codex.acceptance?.requiredTrials !== 9) {
    throw new Error("Codex evidence is not the complete 3 x 3 run");
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
  if (matrix.definitionsSha256 !== codex.definitions?.intentSha256) {
    throw new Error("Matrix and Codex trials did not use the same intent definitions");
  }
  validateMatrixShape(matrix);
  validateCodexShape(codex, matrix.package.tarballSha256);
  validateCandidateMatrixOracles(matrix);
  validateCandidateCodexOracles(codex);
  const companionSummary = validateCompanionShape(companion);
  if (matrix.acceptance.status !== "passed" || matrix.acceptance.passedRuns !== 90 || matrix.acceptance.deterministicIntents !== 30) {
    throw new Error("Candidate matrix did not pass all 90 runs and 30 deterministic intents");
  }
  if (codex.acceptance.status !== "passed" || codex.acceptance.completedTrials !== 9 || codex.acceptance.passedTrials !== 9) {
    throw new Error("Candidate Codex acceptance did not pass all nine fresh-context trials");
  }
  if (companion.acceptance.status !== "passed" || companion.coverage.status !== "passed") {
    throw new Error("Source-derived companion evidence did not pass");
  }
  return {
    schema: "office-kit/pptx-programmable-import-candidate/v1",
    productBaseline: matrix.baseline,
    candidateHead,
    package: {
      name: matrix.package.name,
      version: matrix.package.version,
      tarballSha256: matrix.package.tarballSha256,
      packedCleanInstall: true,
      codexCleanInstallPerTrial: true,
      companionTarballSha256: companion.package.tarballSha256,
    },
    evidenceFiles: {
      matrix: { path: "candidate/matrix.v1.json", sha256: sha256(matrixBytes) },
      codex: { path: "candidate/codex.v1.json", sha256: sha256(codexBytes) },
      companion: { path: "source-derived-companion.evidence.v1.json", sha256: sha256(companionBytes) },
    },
    definitions: {
      matrixSha256: matrix.definitionsSha256,
      continuationSha256: codex.definitions.continuationSha256,
      intentSha256: codex.definitions.intentSha256,
      companionSha256: companion.definitionsSha256,
    },
    matrix: {
      requiredIntents: matrix.acceptance.requiredIntents,
      requiredRuns: matrix.acceptance.requiredRuns,
      passedRuns: matrix.acceptance.passedRuns,
      deterministicIntents: matrix.acceptance.deterministicIntents,
      status: matrix.acceptance.status,
    },
    codex: {
      requiredTasks: codex.acceptance.requiredTasks,
      requiredTrials: codex.acceptance.requiredTrials,
      completedTrials: codex.acceptance.completedTrials,
      passedTrials: codex.acceptance.passedTrials,
      status: codex.acceptance.status,
    },
    sourceDerived: companionSummary,
    environment: {
      matrix: matrix.environment,
      codex: codex.environment,
      companion: companion.environment,
    },
    acceptance: {
      status: "passed",
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

function validateCandidateMatrixOracles(matrix) {
  for (const source of matrix.sources) {
    for (const intent of source.intents) {
      for (const run of intent.runs) {
        const label = `${source.id}/${intent.id}/${run.repetition}`;
        if (run.status !== "passed" || run.sourceSha256After !== source.sourceSha256) throw new Error(`${label}: candidate source/result evidence is invalid`);
        if (run.worker?.sourceUnchanged !== true || run.worker?.secondImport !== true || run.worker?.outputSha256 !== run.outputSha256) {
          throw new Error(`${label}: candidate worker evidence is invalid`);
        }
        if (run.packageOracle?.partSet?.passed !== true || run.packageOracle?.nonTargetPartsByteIdentical !== true
          || run.packageOracle?.relationships?.passed !== true || run.packageOracle?.targetMask?.passed !== true) {
          throw new Error(`${label}: candidate package oracle did not pass`);
        }
        if (run.packageOracle.nestedPackage && run.packageOracle.nestedPackage.passed !== true) {
          throw new Error(`${label}: nested package oracle did not pass`);
        }
        if (run.pixelOracle?.passed !== true || run.pixelOracle?.targetPageChanged !== true
          || run.pixelOracle?.nonTargetPagesPixelIdentical !== true || run.pixelOracle?.nonTargetMismatches?.length !== 0) {
          throw new Error(`${label}: candidate pixel oracle did not pass`);
        }
      }
    }
  }
}

function validateCandidateCodexOracles(codex) {
  const outputByTask = new Map();
  for (const trial of codex.trials) {
    const label = `${trial.taskId}/${trial.repetition}`;
    const checks = trial.checks;
    const commitIds = checks?.durableTask?.commits?.map(({ commitId }) => commitId);
    const outputSha256 = checks?.durableTask?.publication?.sha256;
    if (trial.status !== "passed" || (trial.failures?.length ?? 0) !== 0 || checks?.codex?.passed !== true || checks?.policy?.passed !== true
      || checks?.source?.passed !== true || checks?.output?.passed !== true || checks?.durableTask?.passed !== true
      || checks?.packageOracle?.partSet?.passed !== true || checks?.packageOracle?.relationships?.passed !== true
      || checks?.packageOracle?.targetMask?.passed !== true || checks?.secondImport?.passed !== true
      || checks?.pixelOracle?.passed !== true || checks?.pixelOracle?.nonTargetPagesPixelIdentical !== true) {
      throw new Error(`${label}: candidate Codex oracle set did not pass`);
    }
    if (checks.policy.findings?.length !== 0 || checks.output.createOnly !== true || checks.output.outputCount !== 1) {
      throw new Error(`${label}: candidate Codex policy/output boundary is invalid`);
    }
    if (checks.durableTask.sessions !== 3 || commitIds?.join(",") !== "c0001,c0002"
      || checks.durableTask.head?.commitId !== "c0002" || checks.durableTask.publication?.commitId !== "c0002"
      || checks.durableTask.pending?.length !== 0) {
      throw new Error(`${label}: candidate durable task lifecycle is invalid`);
    }
    if (!outputSha256 || checks.packageOracle.outputSha256 !== outputSha256 || checks.secondImport.inputSha256 !== outputSha256
      || checks.durableTask.head.revisionSha256 !== outputSha256) {
      throw new Error(`${label}: candidate output identity is inconsistent`);
    }
    if (checks.pixelOracle.appendedTargetChangedFromSource !== true || checks.pixelOracle.nonTargetMismatches?.length !== 0) {
      throw new Error(`${label}: candidate visual continuation oracle is invalid`);
    }
    const previousOutput = outputByTask.get(trial.taskId);
    if (previousOutput && previousOutput !== outputSha256) throw new Error(`${trial.taskId}: Codex continuation output is not byte-deterministic`);
    outputByTask.set(trial.taskId, outputSha256);
  }
}

function validateCompanionShape(companion) {
  if (companion.package?.name !== "office-kit" || companion.package?.installKind !== "packed-clean-install" || !companion.package.tarballSha256) {
    throw new Error("Source-derived companion did not use a packed office-kit clean install");
  }
  if (companion.repetitionsPerCase !== 3 || !Array.isArray(companion.cases) || companion.cases.length !== 8) {
    throw new Error("Source-derived companion evidence must contain all eight cases");
  }
  const requiredCoverage = ["text", "geometry", "image", "table", "chart", "component", "add", "delete", "reorder"];
  if (!Array.isArray(companion.coverage?.required) || !Array.isArray(companion.coverage?.passed)) {
    throw new Error("Source-derived companion coverage is missing");
  }
  if (requiredCoverage.some((kind) => !companion.coverage.required.includes(kind) || !companion.coverage.passed.includes(kind))) {
    throw new Error("Source-derived companion does not cover every required operation category");
  }
  let completedRuns = 0;
  let passedRuns = 0;
  const coveredKinds = new Set();
  for (const entry of companion.cases) {
    if (entry.requiredRuns !== 3 || entry.completedRuns !== 3 || entry.passedRuns !== 3 || entry.deterministic !== true) {
      throw new Error(`${entry.id}: source-derived companion case is not 3/3 deterministic`);
    }
    if (!Array.isArray(entry.runs) || entry.runs.length !== 3 || entry.runs.some(({ repetition, status }, index) => repetition !== index + 1 || status !== "passed")) {
      throw new Error(`${entry.id}: source-derived companion runs are invalid`);
    }
    if (new Set(entry.runs.map(({ outputSha256 }) => outputSha256)).size !== 1) {
      throw new Error(`${entry.id}: source-derived companion output is not byte-deterministic`);
    }
    for (const run of entry.runs) {
      const pixelPassed = run.pixelOracle?.passed === true && (
        run.pixelOracle?.pageContentPixelIdentical === true
        || (run.pixelOracle?.targetPageChanged === true && run.pixelOracle?.nonTargetPagesPixelIdentical === true && run.pixelOracle?.nonTargetMismatches?.length === 0)
      );
      if (run.worker?.sourceUnchanged !== true || run.worker?.secondImport?.passed !== true
        || run.packageOracle?.passed !== true || run.packageOracle?.partSet?.passed !== true
        || run.packageOracle?.nonTargetPartsByteIdentical !== true || run.packageOracle?.targetMask?.passed !== true
        || !pixelPassed) {
        throw new Error(`${entry.id}/${run.repetition}: source-derived companion oracle set did not pass`);
      }
    }
    for (const kind of entry.covers || []) coveredKinds.add(kind);
    completedRuns += entry.completedRuns;
    passedRuns += entry.passedRuns;
  }
  for (const kind of companion.existingEvidence?.flatMap(({ passed, covers }) => passed ? covers : []) || []) coveredKinds.add(kind);
  if (requiredCoverage.some((kind) => !coveredKinds.has(kind))) {
    throw new Error("Source-derived companion case evidence does not cover every required operation category");
  }
  return {
    requiredCases: companion.cases.length,
    requiredRuns: companion.cases.length * 3,
    completedRuns,
    passedRuns,
    deterministicCases: companion.cases.filter(({ deterministic }) => deterministic).length,
    coverage: requiredCoverage,
    status: companion.acceptance.status,
  };
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const matrixPath = path.resolve(required(args, "matrix"));
  const codexPath = path.resolve(required(args, "codex"));
  const outputPath = path.resolve(required(args, "output"));
  const [matrixBytes, codexBytes] = await Promise.all([readFile(matrixPath), readFile(codexPath)]);
  const matrix = JSON.parse(matrixBytes.toString("utf8"));
  const codex = JSON.parse(codexBytes.toString("utf8"));
  let evidence;
  if (args.companion) {
    const companionBytes = await readFile(path.resolve(args.companion));
    evidence = buildCandidateEvidence({
      matrix,
      codex,
      companion: JSON.parse(companionBytes.toString("utf8")),
      matrixBytes,
      codexBytes,
      companionBytes,
      candidateHead: required(args, "candidate-head"),
    });
  } else {
    evidence = buildBaselineEvidence({ matrix, codex, matrixBytes, codexBytes, harnessHead: required(args, "harness-head") });
  }
  await writeFile(outputPath, `${JSON.stringify(evidence, null, 2)}\n`, { flag: "wx" });
  process.stdout.write(`${JSON.stringify({ output: outputPath, acceptance: evidence.acceptance, matrix: evidence.matrix.status, codex: evidence.codex.status }, null, 2)}\n`);
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
