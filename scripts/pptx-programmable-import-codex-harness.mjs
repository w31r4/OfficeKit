#!/usr/bin/env node

import { spawn, spawnSync } from "node:child_process";
import { constants } from "node:fs";
import {
  chmod,
  copyFile,
  mkdir,
  readFile,
  readdir,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import {
  compareContinuationRenderedPages,
  evaluateContinuationPackageOracle,
  readContinuationDefinitions,
  readIntentDefinitions,
  renderPresentationPages,
  sha256,
} from "./pptx-programmable-import-oracle.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const defaultContinuationDefinitions = path.join(repoRoot, "evals/pptx-programmable-import/continuation-tasks.v1.json");
const defaultIntentDefinitions = path.join(repoRoot, "evals/pptx-programmable-import/intent-matrix.v1.json");
const DEFAULT_TIMEOUT_MS = 30 * 60 * 1000;
const MAX_CAPTURE_BYTES = 64 * 1024 * 1024;
const TASK_ID_PATTERN = /^t_[a-f0-9]{12}$/u;

export function buildCodexPrompt({ task, source }) {
  const contract = {
    source: "inputs/source.pptx",
    sourceSha256: source.sha256,
    sourceSlides: source.slideCount,
    sourceSlide: task.sourceSlide,
    appendedPage: task.targetPageAfterAppend,
    output: task.output,
    edits: task.edits,
  };
  return `You are one fresh black-box Agent context running an OfficeKit acceptance task.

Read only these installed public workflow files before acting:
- .agents/skills/presentations/SKILL.md
- .agents/skills/office-kit/SKILL.md

Goal:
${task.goal}

Machine contract:
${JSON.stringify(contract, null, 2)}

Use the documented public idioms directly. Import \`FileBlob\`, \`PresentationFile\`, and \`reviewArtifact\` from \`office-kit\` through \`ctx.import\`. Load/import with \`FileBlob.load(staged.path)\` and \`PresentationFile.importPptx(blob)\`; duplicate with \`presentation.slides.items[index].duplicate().moveTo(sourceCount)\`; then call \`PresentationFile.exportPptx\` and reimport before editing the clone. For a native leaf, read records from \`presentation.inspect({ includeNativeLeaves: true, target, maxChars: Infinity }).ndjson.split("\\n").filter(Boolean).map(JSON.parse)\`, require one record matching the contract's target, leaf kind/index, and before value, then call \`presentation.editNativeLeaf(record.targetId, record.leafId, { expectedHash: record.expectedHash, value })\`. For an SVG leaf, resolve the image with \`presentation.resolve(target)\`, require one matching record from \`image.getSvgTextNodes()\`, then call \`image.editSvgText(record.id, { expectedHash: record.expectedHash, value })\`. Do not discover APIs by trial reflection.

Review each bounded edit against the bytes immediately before that edit, not against the original deck with a different page count. In session 1, retain the exported clone-before-edit blob and pass it as \`baseline\` to the imported top-level \`reviewArtifact(outputBlob, options)\`. In session 2, retain the restored c0001 bytes as that edit's \`baseline\`. Pass a not-yet-created \`outputPath\` under \`ctx.taskRoot/candidates\`, set \`layout: false\` and \`visualReview: "unavailable"\`, require a non-failed verdict, and pass the exact output blob plus returned review to \`ctx.commit(outputBlob, { ..., review })\`. Do not save the candidate path before \`reviewArtifact\` writes it.

Complete the workflow in exactly three separate OfficeKit REPL processes inside this one context:
1. Start with \`officekit repl --new <goal> --file phase-1.mjs\`. In that cell call \`ctx.input\` for the immutable source, import it with public \`PresentationFile.importPptx\`, duplicate source slide ${task.sourceSlide}, move the duplicate to the final page, cross an export/import boundary, perform only phase-1's typed edit, call \`reviewArtifact\`, then \`ctx.commit\`. Retain the opaque task id from \`session.ready\`.
2. Start a new process with \`officekit repl <task-id> --file phase-2.mjs\`. Use the restored reviewed full-file revision, not old JavaScript heap state. Reimport it, perform only phase-2's typed edit, call \`reviewArtifact\`, then \`ctx.commit\`.
3. Start a third process with \`officekit repl <task-id> --file phase-3.mjs\`. Reimport the restored current revision, verify the source slide count plus one and both exact values, then publish the current reviewed commit as \`${path.basename(task.output)}\` with \`ctx.publish\`.

Write each phase as one regular UTF-8 \`phase-N.mjs\` cell in the workspace. Run it with \`officekit repl --new <goal> --file phase-1.mjs\` or \`officekit repl <task-id> --file phase-N.mjs\`; the CLI wraps the file into one bounded cell, emits \`session.ready\` and the terminal response, then exits. Do not hand-build JSONL or embed multi-line JavaScript in shell quoting. Parse the first process's \`session.ready\` record to retain the opaque task id, and parse each terminal response before starting the next process.
The \`repl --new\` process itself must receive and execute the phase-1 cell. An empty task-id bootstrap session followed by a fourth REPL process is disqualifying.
Office import/export can run for several minutes. If a shell tool reports that the finite REPL command is still running, keep polling that same process until it exits; never launch a second \`repl --new\`, never duplicate the current phase, and never infer failure from an early partial stdout chunk.

Use artifact id \`continued-deck\` for both commits. Reviews may state visualReview unavailable because the outer independent evaluator owns native rendering, but a failed review must stop the trial. Keep all authored helper files in this workspace and finish only after the exact output path exists.

Hard restrictions:
- Use only the installed public OfficeKit CLI, Skills, and \`office-kit\` package API.
- Follow the documented Skill examples; do not reflect constructors/prototypes or stringify function implementations to discover internals.
- Do not inspect or patch ZIP/OPC/XML/relationship parts. Do not use JSZip, unzip, XPath, or raw package paths.
- Do not use Python, HTML, PPTD, PptxGenJS, \`@oai/artifact-tool\`, or another writer.
- Do not call \`Presentation.create\`, \`slides.add\`, or rebuild the deck. The new page must be the imported source slide's clone.
- Do not edit \`inputs/source.pptx\`, do not write the requested output before the reviewed publication, and do not bypass task/review/commit state by editing \`.office-kit\` files.
- Do not silently retry through a different engine. Preserve any real unsupported or review failure and explain it.

The outer evaluator, not you, performs raw-package, second-import, pixel, source-hash, and clean-install acceptance. Do not create an audit JSON that claims those checks passed.`;
}

export function scanAgentPolicy({ traceText = "", authoredFiles = [] }) {
  const commands = [];
  for (const line of String(traceText).split(/\r?\n/u)) {
    if (!line.trim()) continue;
    try {
      const event = JSON.parse(line);
      if (event?.item?.type === "command_execution" && typeof event.item.command === "string") commands.push(event.item.command);
    } catch {
      // A truncated/non-JSON trace is reported by the Codex exit status. It is
      // not treated as executable text here.
    }
  }
  const findings = [];
  const rules = [
    ["private-runtime", /@oai\/artifact-tool/iu],
    ["python", /(?:^|[\s;&|])(python3?|uv\s+run\s+python)(?:[\s;&|]|$)|\bimport\s+(?:zipfile|lxml|pptx)\b/iu],
    ["html-pptd", /\b(?:pptd|pptxgenjs|PptxGenJS)\b|[.]html?\b/iu],
    ["raw-opc", /\b(?:JSZip|AdmZip|yauzl|unzip|zipinfo|zip|7z|XPath|xml2js|fast-xml-parser)\b|\[Content_Types\]|<p:|[.]rels\b|ppt\/slides\/slide\d+[.]xml/iu],
    ["whole-rebuild", /\bPresentation[.]create\s*\(|[.]slides[.]add\s*\(/u],
    ["task-store-bypass", /[.]office-kit\/tasks|task[.]json/iu],
    ["repository-internal", /(?:\.\.\/)+(?:src|native|runtime)\/|\/src\/index[.]mjs/iu],
    ["api-reflection", /Object[.]getOwnPropertyNames\s*\(|\b(?:FileBlob|PresentationFile)[.]toString\s*\(/u],
  ];
  const sources = [
    ...commands.map((text, index) => ({ origin: `trace-command:${index + 1}`, text })),
    ...authoredFiles,
  ];
  for (const source of sources) {
    for (const [code, pattern] of rules) {
      if (pattern.test(source.text)) findings.push({ code, origin: source.origin });
    }
  }
  return {
    passed: findings.length === 0,
    commandCount: commands.length,
    authoredFileCount: authoredFiles.length,
    findings,
  };
}

export async function evaluateDurableTask({ workspace, task, source, outputPath }) {
  const taskStore = path.join(workspace, ".office-kit", "tasks");
  const entries = await readdir(taskStore, { withFileTypes: true });
  const taskIds = entries.filter((entry) => entry.isDirectory() && TASK_ID_PATTERN.test(entry.name)).map(({ name }) => name);
  if (taskIds.length !== 1) throw new Error(`${task.id}: expected one durable task, observed ${taskIds.length}`);
  const taskId = taskIds[0];
  const taskRoot = path.join(taskStore, taskId);
  const manifest = JSON.parse(await readFile(path.join(taskRoot, "task.json"), "utf8"));
  if (manifest.id !== taskId || manifest.schemaVersion !== 1) throw new Error(`${task.id}: invalid durable task manifest identity`);
  if (manifest.commits?.length !== 2 || manifest.commits[0]?.id !== "c0001" || manifest.commits[1]?.id !== "c0002") {
    throw new Error(`${task.id}: expected exactly c0001 and c0002`);
  }
  if (manifest.head?.commitId !== "c0002") throw new Error(`${task.id}: durable HEAD is not c0002`);
  if (manifest.pending?.length !== 0) throw new Error(`${task.id}: durable task has ${manifest.pending.length} unresolved pending records`);
  if (manifest.publications?.length !== 1) throw new Error(`${task.id}: expected one publication, observed ${manifest.publications?.length ?? 0}`);
  const sessions = (await readdir(path.join(taskRoot, "sessions"), { withFileTypes: true })).filter((entry) => entry.isDirectory());
  if (sessions.length !== 3) throw new Error(`${task.id}: expected three REPL sessions, observed ${sessions.length}`);
  if (new Set(sessions.map(({ name }) => name)).size !== 3 || !sessions.some(({ name }) => name === manifest.lastSessionId)) {
    throw new Error(`${task.id}: REPL session identity is inconsistent`);
  }
  const sourceArtifact = manifest.artifacts?.find(({ source: artifactSource }) => artifactSource);
  const continuedArtifact = manifest.artifacts?.find(({ id }) => id === "continued-deck");
  if (!sourceArtifact || sourceArtifact.source.sha256 !== source.sha256) throw new Error(`${task.id}: staged source provenance is missing or stale`);
  if (!continuedArtifact || continuedArtifact.kind !== "presentation") throw new Error(`${task.id}: continued-deck artifact is missing`);

  const revisions = [];
  for (const commit of manifest.commits) {
    if (commit.artifactId !== "continued-deck") throw new Error(`${task.id}: commit ${commit.id} targets ${commit.artifactId}`);
    if (!new Set(["passed", "passed-with-limitations"]).has(commit.review?.verdict)) throw new Error(`${task.id}: commit ${commit.id} is not reviewed`);
    if (commit.review.deliverySha256 !== commit.revisionSha256) throw new Error(`${task.id}: commit ${commit.id} review hash is stale`);
    const evidence = commit.review.evidence;
    const evidencePath = resolveManaged(taskRoot, evidence?.path, "review evidence");
    const evidenceBytes = await readFile(evidencePath);
    if (sha256(evidenceBytes) !== evidence.sha256) throw new Error(`${task.id}: commit ${commit.id} review evidence hash mismatch`);
    const revision = commit.heads?.["continued-deck"];
    const revisionPath = resolveManaged(taskRoot, revision?.path, "revision");
    const bytes = await readFile(revisionPath);
    if (sha256(bytes) !== revision.sha256 || revision.sha256 !== commit.revisionSha256) throw new Error(`${task.id}: commit ${commit.id} revision hash mismatch`);
    revisions.push({ commitId: commit.id, path: revisionPath, sha256: revision.sha256, bytes: bytes.byteLength, reviewVerdict: commit.review.verdict });
  }
  const outputBytes = await readFile(outputPath);
  const publication = manifest.publications[0];
  if (path.resolve(publication.path) !== path.resolve(outputPath)) throw new Error(`${task.id}: publication path is not the required output`);
  if (publication.commitId !== "c0002" || publication.artifactId !== "continued-deck") throw new Error(`${task.id}: publication is not current reviewed HEAD`);
  if (sha256(outputBytes) !== publication.sha256 || publication.sha256 !== revisions[1].sha256) throw new Error(`${task.id}: publication bytes are not c0002`);
  return {
    passed: true,
    taskId,
    schemaVersion: manifest.schemaVersion,
    sessions: sessions.length,
    commits: revisions,
    head: manifest.head,
    pending: manifest.pending,
    publication: { commitId: publication.commitId, artifactId: publication.artifactId, path: path.relative(workspace, publication.path), sha256: publication.sha256 },
    sourceProvenance: { artifactId: sourceArtifact.id, sha256: sourceArtifact.source.sha256, storedPath: sourceArtifact.source.storedPath },
  };
}

export async function inspectAgentRuntimeState(workspace) {
  const taskStore = path.join(workspace, ".office-kit", "tasks");
  let entries;
  try {
    entries = await readdir(taskStore, { withFileTypes: true });
  } catch (error) {
    if (error.code === "ENOENT") return null;
    throw error;
  }
  const taskIds = entries.filter((entry) => entry.isDirectory() && TASK_ID_PATTERN.test(entry.name)).map(({ name }) => name);
  if (taskIds.length > 1) return { code: "multiple-durable-tasks", observed: taskIds.length, maximum: 1 };
  let sessions = 0;
  for (const taskId of taskIds) {
    const sessionsRoot = path.join(taskStore, taskId, "sessions");
    sessions += (await readdir(sessionsRoot, { withFileTypes: true }).catch(() => [])).filter((entry) => entry.isDirectory()).length;
  }
  if (sessions > 3) return { code: "repl-session-budget-exceeded", observed: sessions, maximum: 3 };
  return null;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const continuationPath = path.resolve(args.definitions || defaultContinuationDefinitions);
  const intentsPath = path.resolve(args["intent-definitions"] || defaultIntentDefinitions);
  const [definitions, intents] = await Promise.all([readContinuationDefinitions(continuationPath), readIntentDefinitions(intentsPath)]);
  if (definitions.baseline !== intents.baseline) throw new Error("Continuation and intent baselines differ");
  const assetsDir = path.resolve(args["assets-dir"] || process.env[args["assets-environment"] || intents.assetsEnvironment] || intents.defaultAssetsDirectory);
  const runRoot = path.resolve(required(args, "run-root"));
  await requireAbsent(runRoot, "run root");
  await mkdir(runRoot, { recursive: true });
  const pack = await packCandidate(runRoot);
  const render = args["no-render"] !== true;
  const repetitions = args.trials ? positiveInteger(args.trials, "trials") : definitions.trialsPerTask;
  const selectedTasks = args.task ? definitions.tasks.filter(({ id }) => id === args.task) : definitions.tasks;
  if (!selectedTasks.length) throw new Error(`Unknown --task ${args.task}`);
  const renderCache = path.join(runRoot, "render-cache");
  await mkdir(renderCache);
  const trials = [];
  for (const task of selectedTasks) {
    const source = intents.sources.find(({ id }) => id === task.sourceId);
    if (!source) throw new Error(`${task.id}: unknown source ${task.sourceId}`);
    const sourcePath = path.join(assetsDir, source.fileName);
    const sourceBytes = await readFile(sourcePath);
    if (sha256(sourceBytes) !== source.sha256) throw new Error(`${source.id}: source SHA-256 mismatch`);
    const sourceRender = render ? await renderPresentationPages(sourcePath, renderCache, source.sha256) : null;
    for (let repetition = 1; repetition <= repetitions; repetition += 1) {
      trials.push(await runTrial({
        runRoot,
        pack,
        definitionsPath: continuationPath,
        task,
        source,
        sourcePath,
        sourceBytes,
        sourceRender,
        renderCache,
        render,
        repetition,
        timeoutMs: args["timeout-ms"] ? positiveInteger(args["timeout-ms"], "timeout-ms", 24 * 60 * 60 * 1000) : Number(process.env.OFFICEKIT_PPTX_CODEX_TIMEOUT_MS || DEFAULT_TIMEOUT_MS),
      }));
    }
  }
  const evidence = {
    schema: "office-kit/pptx-codex-continuation-evidence/v1",
    baseline: definitions.baseline,
    definitions: {
      continuationSha256: sha256(await readFile(continuationPath)),
      intentSha256: sha256(await readFile(intentsPath)),
    },
    package: Object.fromEntries(Object.entries(pack).filter(([key]) => key !== "tarballPath")),
    environment: {
      platform: process.platform,
      arch: process.arch,
      node: process.version,
      npm: versionLine(process.platform === "win32" ? "npm.cmd" : "npm", ["--version"]),
      codex: versionLine(process.env.OFFICEKIT_CODEX_BIN || "codex", ["--version"]),
      render,
    },
    protocol: {
      freshCodexContextPerTrial: true,
      replSessionsPerTrial: definitions.replSessionsPerTrial,
      publicPackageOnly: true,
      createOnlyOutputs: true,
    },
    trials,
    acceptance: {
      requiredTasks: selectedTasks.length,
      trialsPerTask: repetitions,
      requiredTrials: selectedTasks.length * repetitions,
      completedTrials: trials.length,
      passedTrials: trials.filter(({ status }) => status === "passed").length,
      status: trials.length === selectedTasks.length * repetitions && trials.every(({ status }) => status === "passed") ? "passed" : "failed",
    },
  };
  const evidencePath = path.join(runRoot, "evidence.json");
  await writeFile(evidencePath, `${JSON.stringify(evidence, null, 2)}\n`, { flag: "wx" });
  process.stdout.write(`${JSON.stringify({ evidence: evidencePath, acceptance: evidence.acceptance }, null, 2)}\n`);
  if (evidence.acceptance.status !== "passed") process.exitCode = 1;
}

async function runTrial({ runRoot, pack, definitionsPath, task, source, sourcePath, sourceBytes, sourceRender, renderCache, render, repetition, timeoutMs }) {
  const trialRoot = path.join(runRoot, "trials", task.id, String(repetition));
  const workspace = path.join(trialRoot, "workspace");
  const evaluator = path.join(trialRoot, "evaluator");
  await mkdir(path.join(workspace, "inputs"), { recursive: true });
  await mkdir(evaluator, { recursive: true });
  const inputPath = path.join(workspace, "inputs/source.pptx");
  await copyFile(sourcePath, inputPath, constants.COPYFILE_EXCL);
  await chmod(inputPath, 0o444);
  await writeFile(path.join(workspace, "package.json"), `${JSON.stringify({ name: `officekit-pptx-acceptance-${task.id}-${repetition}`, private: true }, null, 2)}\n`, { flag: "wx" });
  const install = runRequired(process.platform === "win32" ? "npm.cmd" : "npm", ["install", "--ignore-scripts", "--no-audit", "--no-fund", pack.tarballPath], workspace, `${task.id}/${repetition}: packed clean install`);
  await writeFile(path.join(evaluator, "npm-install.stdout.txt"), install.stdout, { flag: "wx" });
  await writeFile(path.join(evaluator, "npm-install.stderr.txt"), install.stderr, { flag: "wx" });
  const officekitBin = path.join(workspace, "node_modules/office-kit/bin/officekit.mjs");
  const installedMetadata = JSON.parse(await readFile(path.join(workspace, "node_modules/office-kit/package.json"), "utf8"));
  if (installedMetadata.name !== "office-kit" || installedMetadata.version !== pack.version) throw new Error(`${task.id}/${repetition}: clean install identity mismatch`);
  const initialized = runRequired(process.execPath, [officekitBin, "init", "--tools", "agents", "--yes", "--json"], workspace, `${task.id}/${repetition}: officekit init`);
  await writeFile(path.join(evaluator, "officekit-init.json"), initialized.stdout, { flag: "wx" });
  await retainAcceptanceSkills(workspace);
  const prompt = buildCodexPrompt({ task, source });
  await writeFile(path.join(evaluator, "prompt.md"), `${prompt}\n`, { flag: "wx" });
  await writeFile(path.join(evaluator, "prompt.sha256"), `${sha256(Buffer.from(prompt))}\n`, { flag: "wx" });
  const codex = await runCodex({ workspace, evaluator, prompt, timeoutMs });
  const { traceText, ...codexRecord } = codex;
  const outputPath = path.join(workspace, task.output);
  const failures = [];
  const checks = {};
  checks.codex = { passed: codex.status === 0, ...codexRecord };
  if (codex.policyTermination) failures.push(`Codex policy termination: ${codex.policyTermination.code} (${codex.policyTermination.observed} > ${codex.policyTermination.maximum})`);
  else if (codex.status !== 0) failures.push(`Codex exited ${codex.status}${codex.timedOut ? " after timeout" : ""}`);
  const authoredFiles = [...await readAuthoredFiles(workspace), ...await readReplSources(workspace)];
  checks.policy = scanAgentPolicy({ traceText: codex.traceText, authoredFiles });
  if (!checks.policy.passed) failures.push(`Agent policy violations: ${checks.policy.findings.map(({ code, origin }) => `${code}@${origin}`).join(", ")}`);
  const sourceAfter = sha256(await readFile(inputPath));
  checks.source = { passed: sourceAfter === source.sha256, expectedSha256: source.sha256, afterSha256: sourceAfter, immutableMode: (await stat(inputPath)).mode & 0o777 };
  if (!checks.source.passed) failures.push(`Source SHA changed: ${sourceAfter}`);
  try {
    await stat(outputPath);
    const outputPptx = (await readdir(path.join(workspace, "outputs"), { withFileTypes: true })).filter((entry) => entry.isFile() && entry.name.endsWith(".pptx"));
    if (outputPptx.length !== 1 || outputPptx[0].name !== path.basename(task.output)) throw new Error(`expected one exact output PPTX, observed ${outputPptx.map(({ name }) => name).join(", ")}`);
    checks.output = { passed: true, path: task.output, createOnly: true, outputCount: 1 };
  } catch (error) {
    checks.output = { passed: false, reason: errorMessage(error) };
    failures.push(`Output contract: ${checks.output.reason}`);
  }
  if (checks.output.passed) {
    const outputBytes = await readFile(outputPath);
    try {
      checks.durableTask = await evaluateDurableTask({ workspace, task, source, outputPath });
    } catch (error) {
      checks.durableTask = { passed: false, reason: errorMessage(error) };
      failures.push(`Durable task: ${checks.durableTask.reason}`);
    }
    try {
      checks.packageOracle = await evaluateContinuationPackageOracle({ sourceBytes, outputBytes, source, task });
    } catch (error) {
      checks.packageOracle = { passed: false, reason: errorMessage(error) };
      failures.push(`Package oracle: ${checks.packageOracle.reason}`);
    }
    try {
      const receiptPath = path.join(evaluator, "second-import.json");
      const verifier = runRequired(process.execPath, [
        officekitBin,
        "run",
        path.join(repoRoot, "scripts/pptx-programmable-import-continuation-verify.mjs"),
        "--definitions", definitionsPath,
        "--task-id", task.id,
        "--input", outputPath,
        "--receipt", receiptPath,
      ], workspace, `${task.id}/${repetition}: packed public second import`);
      await writeFile(path.join(evaluator, "second-import.stdout.txt"), verifier.stdout, { flag: "wx" });
      await writeFile(path.join(evaluator, "second-import.stderr.txt"), verifier.stderr, { flag: "wx" });
      checks.secondImport = JSON.parse(await readFile(receiptPath, "utf8"));
      checks.secondImport.passed = true;
    } catch (error) {
      checks.secondImport = { passed: false, reason: errorMessage(error) };
      failures.push(`Second import: ${checks.secondImport.reason}`);
    }
    if (render) {
      try {
        const outputRender = await renderPresentationPages(outputPath, renderCache, sha256(outputBytes));
        checks.pixelOracle = compareContinuationRenderedPages(sourceRender, outputRender, task.targetPageAfterAppend);
      } catch (error) {
        checks.pixelOracle = { passed: false, reason: errorMessage(error) };
        failures.push(`Pixel oracle: ${checks.pixelOracle.reason}`);
      }
    } else {
      checks.pixelOracle = { passed: false, skipped: true, reason: "--no-render" };
    }
  }
  const replacements = [[workspace, "$WORKSPACE"], [runRoot, "$RUN_ROOT"], [repoRoot, "$REPOSITORY"]];
  return {
    taskId: task.id,
    sourceId: task.sourceId,
    repetition,
    freshCodexContext: true,
    packedCleanInstall: { passed: true, package: installedMetadata.name, version: installedMetadata.version, tarballSha256: pack.tarballSha256 },
    status: failures.length === 0 ? "passed" : "failed",
    failures: portableize(failures, replacements),
    checks: portableize(checks, replacements),
    evidenceDirectory: path.relative(runRoot, evaluator),
  };
}

async function packCandidate(runRoot) {
  const packRoot = path.join(runRoot, "package");
  await mkdir(packRoot);
  const result = runRequired(process.platform === "win32" ? "npm.cmd" : "npm", ["pack", "--json", "--ignore-scripts", "--pack-destination", packRoot], repoRoot, "npm pack candidate");
  const records = JSON.parse(result.stdout);
  if (!Array.isArray(records) || records.length !== 1 || !records[0].filename) throw new Error("npm pack did not return one tarball");
  const tarballPath = path.join(packRoot, records[0].filename);
  const tarballBytes = await readFile(tarballPath);
  const metadata = JSON.parse(await readFile(path.join(repoRoot, "package.json"), "utf8"));
  return {
    name: metadata.name,
    version: metadata.version,
    baselineCandidate: true,
    tarballPath,
    tarballFile: records[0].filename,
    tarballSha256: sha256(tarballBytes),
    packedBytes: records[0].size,
    unpackedBytes: records[0].unpackedSize,
    totalFiles: records[0].entryCount,
    cleanInstallPerTrial: true,
    lifecycleScripts: "ignored",
  };
}

async function runCodex({ workspace, evaluator, prompt, timeoutMs }) {
  const tracePath = path.join(evaluator, "codex-trace.jsonl");
  const stderrPath = path.join(evaluator, "codex-stderr.txt");
  const finalPath = path.join(evaluator, "codex-final.txt");
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
    ...(process.env.OFFICEKIT_PPTX_CODEX_MODEL ? ["--model", process.env.OFFICEKIT_PPTX_CODEX_MODEL] : []),
    "-o", finalPath,
    "-",
  ];
  const child = spawn(process.env.OFFICEKIT_CODEX_BIN || "codex", args, {
    cwd: workspace,
    detached: process.platform !== "win32",
    env: {
      ...process.env,
      PATH: `${path.join(workspace, "node_modules/.bin")}${path.delimiter}${process.env.PATH || ""}`,
      PYTHONDONTWRITEBYTECODE: "1",
    },
    stdio: ["pipe", "pipe", "pipe"],
  });
  let traceText = "";
  let stderr = "";
  let captureExceeded = false;
  child.stdout.setEncoding("utf8");
  child.stderr.setEncoding("utf8");
  child.stdout.on("data", (chunk) => {
    if (Buffer.byteLength(traceText) + Buffer.byteLength(chunk) > MAX_CAPTURE_BYTES) captureExceeded = true;
    else traceText += chunk;
  });
  child.stderr.on("data", (chunk) => {
    if (Buffer.byteLength(stderr) + Buffer.byteLength(chunk) > MAX_CAPTURE_BYTES) captureExceeded = true;
    else stderr += chunk;
  });
  child.stdin.end(prompt);
  let interruptedSignal = null;
  const onInterrupt = (signal) => {
    interruptedSignal = signal;
    terminateProcessTree(child.pid);
  };
  process.once("SIGINT", onInterrupt);
  process.once("SIGTERM", onInterrupt);
  let timedOut = false;
  let policyTermination = null;
  let monitorBusy = false;
  const policyMonitor = setInterval(async () => {
    if (monitorBusy || policyTermination) return;
    monitorBusy = true;
    try {
      const violation = await inspectAgentRuntimeState(workspace);
      if (violation) {
        policyTermination = violation;
        terminateProcessTree(child.pid);
        setTimeout(() => terminateProcessTree(child.pid, "SIGKILL"), 5000).unref();
      }
    } catch {
      // The final evaluator re-reads the task store. A transient read race must
      // not invent a live policy failure.
    } finally {
      monitorBusy = false;
    }
  }, 1000);
  policyMonitor.unref();
  const timer = setTimeout(() => {
    timedOut = true;
    terminateProcessTree(child.pid);
    setTimeout(() => terminateProcessTree(child.pid, "SIGKILL"), 5000).unref();
  }, timeoutMs);
  const result = await new Promise((resolve) => {
    let settled = false;
    const finish = (value) => {
      if (settled) return;
      settled = true;
      resolve(value);
    };
    child.once("close", (status, signal) => finish({ status, signal }));
    child.once("error", (error) => finish({ status: 127, signal: null, error: error.message }));
    setTimeout(() => {
      if (timedOut) finish({ status: 124, signal: "SIGTERM" });
    }, timeoutMs + 7000).unref();
  });
  clearTimeout(timer);
  clearInterval(policyMonitor);
  process.off("SIGINT", onInterrupt);
  process.off("SIGTERM", onInterrupt);
  if (captureExceeded) terminateProcessTree(child.pid);
  const status = policyTermination ? 126 : interruptedSignal === "SIGINT" ? 130 : interruptedSignal === "SIGTERM" ? 143 : captureExceeded ? 125 : timedOut ? 124 : (result.status ?? 1);
  await writeFile(tracePath, traceText, { flag: "wx" });
  await writeFile(stderrPath, stderr, { flag: "wx" });
  await writeFile(path.join(evaluator, "codex-exit.json"), `${JSON.stringify({ status, signal: result.signal, timedOut, captureExceeded, timeoutMs, error: result.error || null }, null, 2)}\n`, { flag: "wx" });
  const finalBytes = await readFile(finalPath).catch(() => null);
  const finalLimit = 16 * 1024;
  return {
    status,
    signal: result.signal,
    timedOut,
    interruptedSignal,
    policyTermination,
    captureExceeded,
    timeoutMs,
    traceSha256: sha256(Buffer.from(traceText)),
    stderrSha256: sha256(Buffer.from(stderr)),
    finalSha256: finalBytes ? sha256(finalBytes) : null,
    agentFinal: finalBytes ? finalBytes.subarray(0, finalLimit).toString("utf8") : null,
    agentFinalTruncated: Boolean(finalBytes && finalBytes.byteLength > finalLimit),
    traceText,
  };
}

function terminateProcessTree(pid, signal = "SIGTERM") {
  const descendants = descendantProcessIds(pid);
  for (const childPid of descendants.reverse()) {
    try { process.kill(childPid, signal); } catch {}
    try { if (process.platform !== "win32") process.kill(-childPid, signal); } catch {}
  }
  try {
    if (process.platform !== "win32") process.kill(-pid, signal);
    else process.kill(pid, signal);
  } catch {}
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

async function retainAcceptanceSkills(workspace) {
  const skillsRoot = path.join(workspace, ".agents", "skills");
  const entries = await readdir(skillsRoot, { withFileTypes: true });
  for (const entry of entries) {
    if (entry.isDirectory() && !new Set(["presentations", "office-kit"]).has(entry.name)) {
      await rm(path.join(skillsRoot, entry.name), { recursive: true, force: true });
    }
  }
  for (const required of ["presentations", "office-kit"]) await stat(path.join(skillsRoot, required, "SKILL.md"));
}

async function readAuthoredFiles(workspace) {
  const output = [];
  async function walk(directory, relative = "") {
    for (const entry of await readdir(directory, { withFileTypes: true })) {
      const childRelative = path.posix.join(relative, entry.name);
      if (!relative && new Set(["node_modules", ".agents", ".office-kit", "inputs", "outputs", ".git"]).has(entry.name)) continue;
      const absolute = path.join(directory, entry.name);
      if (entry.isDirectory()) {
        await walk(absolute, childRelative);
        continue;
      }
      if (!entry.isFile() || new Set(["package.json", "package-lock.json"]).has(childRelative)) continue;
      const extension = path.extname(entry.name).toLowerCase();
      if (new Set([".py", ".html", ".htm", ".pptd", ".xml"]).has(extension)) {
        output.push({ origin: `workspace:${childRelative}`, text: `${extension} authored file` });
      } else if (new Set([".js", ".mjs", ".cjs", ".sh", ".md", ".txt"]).has(extension)) {
        output.push({ origin: `workspace:${childRelative}`, text: await readFile(absolute, "utf8") });
      }
    }
  }
  await walk(workspace);
  return output;
}

async function readReplSources(workspace) {
  const taskStore = path.join(workspace, ".office-kit", "tasks");
  const output = [];
  let taskEntries;
  try {
    taskEntries = await readdir(taskStore, { withFileTypes: true });
  } catch (error) {
    if (error.code === "ENOENT") return output;
    throw error;
  }
  for (const taskEntry of taskEntries) {
    if (!taskEntry.isDirectory() || !TASK_ID_PATTERN.test(taskEntry.name)) continue;
    const sessionsRoot = path.join(taskStore, taskEntry.name, "sessions");
    for (const sessionEntry of await readdir(sessionsRoot, { withFileTypes: true }).catch(() => [])) {
      if (!sessionEntry.isDirectory()) continue;
      const journal = await readFile(path.join(sessionsRoot, sessionEntry.name, "session.jsonl"), "utf8").catch(() => "");
      for (const line of journal.split(/\r?\n/u)) {
        if (!line.trim()) continue;
        try {
          const record = JSON.parse(line);
          if (record.type === "request.started" && typeof record.source === "string") {
            output.push({ origin: `repl:${taskEntry.name}/${sessionEntry.name}/${record.id ?? record.sequence}`, text: record.source });
          }
        } catch {}
      }
    }
  }
  return output;
}

function resolveManaged(root, relative, label) {
  if (typeof relative !== "string" || path.isAbsolute(relative) || relative.split(/[\\/]/u).includes("..")) throw new Error(`Unsafe ${label} path`);
  const resolved = path.resolve(root, relative);
  if (resolved !== root && !resolved.startsWith(`${root}${path.sep}`)) throw new Error(`Escaped ${label} path`);
  return resolved;
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

function positiveInteger(value, name, maximum = 10) {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 1 || parsed > maximum) throw new Error(`${name} must be an integer from 1 through ${maximum}`);
  return parsed;
}

async function requireAbsent(target, label) {
  try {
    await stat(target);
  } catch (error) {
    if (error.code === "ENOENT") return;
    throw error;
  }
  throw new Error(`${label} already exists; evidence is create-only: ${target}`);
}

function errorMessage(error) {
  return error instanceof Error ? error.message : String(error);
}

function portableize(value, replacements) {
  let encoded = JSON.stringify(value);
  for (const [absolute, token] of replacements.sort(([left], [right]) => right.length - left.length)) {
    encoded = encoded.replaceAll(absolute, token);
  }
  return JSON.parse(encoded);
}

if (import.meta.url === pathToFileURL(process.argv[1] || "").href) {
  await main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
