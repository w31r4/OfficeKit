import { createHash, randomUUID } from "node:crypto";
import {
  chmod,
  lstat,
  mkdir,
  open,
  readFile,
  readdir,
  realpath,
  rename,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import {
  MAX_AUTHORING_PLAN_BYTES,
  PRESENTATION_AUTHORING_PLAN_SCHEMA,
  PRESENTATION_COMMUNICATION_JOBS,
  PRESENTATION_MEDIUM_FITS,
  PRESENTATION_SCENARIOS,
  authoringPlanDescriptor,
  normalizePresentationAuthoringPlan,
} from "./authoring-plan.mjs";

export const TASK_SCHEMA_VERSION = 2;
export const LEGACY_TASK_SCHEMA_VERSION = 1;
export const DEFAULT_TASK_LIST_LIMIT = 5;
export const DEFAULT_MAX_TASK_MANIFEST_BYTES = 1_048_576;
export const DEFAULT_MAX_TASK_ARTIFACT_BYTES = 536_870_912;
export const DEFAULT_MAX_REVIEW_REPORT_BYTES = 8_388_608;
export const DEFAULT_MAX_TASK_OPERATION_BYTES = 1_048_576;
export const DEFAULT_MAX_TASK_PPJ_BYTES = 16_777_216;

const TASK_ID_PATTERN = /^t_[a-f0-9]{12}$/u;
const ARTIFACT_ID_PATTERN = /^[a-z0-9][a-z0-9._-]{0,63}$/u;
const TASK_DIRECTORY = path.join(".office-kit", "tasks");
const TASK_IGNORE = "*\n!.gitignore\n";
const ARTIFACT_KINDS = new Set(["document", "workbook", "presentation", "pdf"]);
const VISUAL_REVIEW_STATUSES = new Set(["complete", "unavailable", "requires-human"]);
const RESOLVED_BY_COMMIT = new Set(["review-failed", "stale-review"]);
const PRESENTATION_COMMUNICATION_JOB_SET = new Set(PRESENTATION_COMMUNICATION_JOBS);
const PRESENTATION_MEDIUM_FIT_SET = new Set(PRESENTATION_MEDIUM_FITS);
const PRESENTATION_SCENARIO_SET = new Set(PRESENTATION_SCENARIOS);
const PRESENTATION_STRATEGY_STATUS_SET = new Set(["current", "legacy"]);
const PPJ_TASK_SCHEMA = "office-kit/ppj-task/v1";
const PPJ_PROGRAM_SCHEMA = "office-kit/ppj/v1";
const PPJ_REVISION_STATUSES = new Set(["valid", "candidate", "reviewed", "review-failed"]);
const PPJ_RECEIPT_STAGES = new Set(["imported", "checked", "built", "reviewed"]);
const MAX_PPJ_REVISIONS = 256;

export async function resolveTaskWorkspace({ workspaceRoot, cwd = process.cwd() } = {}) {
  if (workspaceRoot != null) return canonicalDirectory(workspaceRoot, "workspace");
  let current = await canonicalDirectory(cwd, "workspace");
  let gitRoot;
  while (true) {
    const officeKit = path.join(current, ".office-kit");
    const officeKitStat = await lstatIfExists(officeKit);
    if (officeKitStat) {
      if (officeKitStat.isSymbolicLink() || !officeKitStat.isDirectory()) {
        throw taskError("unsafe-workspace", `.office-kit must be a regular directory: ${officeKit}`);
      }
      return current;
    }
    if (gitRoot == null) {
      const gitMarker = await lstatIfExists(path.join(current, ".git"));
      if (gitMarker && !gitMarker.isSymbolicLink() && (gitMarker.isDirectory() || gitMarker.isFile())) {
        gitRoot = current;
      }
    }
    const parent = path.dirname(current);
    if (parent === current) break;
    current = parent;
  }
  return gitRoot ?? canonicalDirectory(cwd, "workspace");
}

export async function createTask({ workspaceRoot, goal, now = new Date() }) {
  const workspace = await resolveTaskWorkspace({ workspaceRoot });
  const normalizedGoal = boundedText(goal, "Task goal", 1_024);
  const tasksRoot = await ensureTaskStore(workspace);
  let taskId;
  let taskRoot;
  for (let attempt = 0; attempt < 8; attempt += 1) {
    taskId = `t_${randomUUID().replaceAll("-", "").slice(0, 12)}`;
    taskRoot = path.join(tasksRoot, taskId);
    try {
      await mkdir(taskRoot, { mode: 0o700 });
      break;
    } catch (error) {
      if (error?.code !== "EEXIST" || attempt === 7) throw error;
    }
  }
  await privateMode(taskRoot, 0o700);
  for (const directory of ["inputs", "revisions", "candidates", "evidence", "operations", "plans", "programs", "sessions"]) {
    await ensurePrivateDirectory(path.join(taskRoot, directory), taskRoot);
  }
  const timestamp = now.toISOString();
  const manifest = {
    schemaVersion: TASK_SCHEMA_VERSION,
    id: taskId,
    goal: normalizedGoal,
    createdAt: timestamp,
    updatedAt: timestamp,
    next: null,
    constraints: [],
    artifacts: [],
    commits: [],
    head: null,
    pending: [],
    publications: [],
    lastSessionId: null,
    plan: null,
    ppj: null,
  };
  await writeTaskManifest(taskRoot, manifest);
  return { workspaceRoot: workspace, taskRoot, manifest };
}

export async function openTask({ workspaceRoot, taskId }) {
  const workspace = await resolveTaskWorkspace({ workspaceRoot });
  const taskRoot = await resolveTaskRoot(workspace, taskId);
  const manifest = await readTaskManifest(taskRoot, taskId);
  return { workspaceRoot: workspace, taskRoot, manifest };
}

export async function listTasks({ workspaceRoot, all = false, limit = DEFAULT_TASK_LIST_LIMIT } = {}) {
  const workspace = await resolveTaskWorkspace({ workspaceRoot });
  const tasksRoot = path.join(workspace, TASK_DIRECTORY);
  const tasksRootStat = await lstatIfExists(tasksRoot);
  if (tasksRootStat == null) return { workspace, total: 0, shown: 0, truncated: false, tasks: [], invalid: [] };
  if (tasksRootStat.isSymbolicLink() || !tasksRootStat.isDirectory()) {
    throw taskError("unsafe-task-store", `Task store must be a regular directory: ${tasksRoot}`);
  }
  const canonicalTasksRoot = await realpath(tasksRoot);
  const entries = await readdir(canonicalTasksRoot, { withFileTypes: true });
  const tasks = [];
  const invalid = [];
  for (const entry of entries.sort((a, b) => a.name.localeCompare(b.name))) {
    if (entry.name === ".gitignore") continue;
    if (!TASK_ID_PATTERN.test(entry.name) || !entry.isDirectory() || entry.isSymbolicLink()) {
      invalid.push({ id: entry.name, code: "invalid-task-entry" });
      continue;
    }
    try {
      const taskRoot = path.join(canonicalTasksRoot, entry.name);
      const manifest = await readTaskManifest(taskRoot, entry.name);
      tasks.push(summarizeTask(manifest, { detailed: false }));
    } catch (error) {
      invalid.push({ id: entry.name, code: error?.code || "invalid-task", message: boundedError(error) });
    }
  }
  tasks.sort((left, right) => right.updatedAt.localeCompare(left.updatedAt) || left.id.localeCompare(right.id));
  const visible = all ? tasks : tasks.slice(0, normalizeListLimit(limit));
  return {
    workspace,
    total: tasks.length,
    shown: visible.length,
    truncated: visible.length < tasks.length,
    tasks: visible,
    invalid,
  };
}

export async function taskDetail({ workspaceRoot, taskId }) {
  const opened = await openTask({ workspaceRoot, taskId });
  const summary = summarizeTask(opened.manifest, { detailed: true, taskRoot: opened.taskRoot });
  summary.storageBytes = await directoryBytes(opened.taskRoot);
  return {
    workspace: opened.workspaceRoot,
    task: summary,
  };
}

export async function deleteTask({ workspaceRoot, taskId }) {
  const opened = await openTask({ workspaceRoot, taskId });
  const lock = await lstatIfExists(path.join(opened.taskRoot, ".write.lock"));
  if (lock) throw taskError("task-busy", `Task ${taskId} is open and cannot be deleted.`);
  const bytes = await directoryBytes(opened.taskRoot);
  await rm(opened.taskRoot, { recursive: true, force: false });
  return { workspace: opened.workspaceRoot, taskId, deleted: true, bytes };
}

export async function acquireTaskLock(taskRoot, { sessionId, now = new Date() }) {
  const target = path.join(taskRoot, ".write.lock");
  const token = randomUUID();
  const record = { schemaVersion: 1, pid: process.pid, sessionId, token, createdAt: now.toISOString() };
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      const handle = await open(target, "wx", 0o600);
      try { await handle.writeFile(`${JSON.stringify(record)}\n`, "utf8"); }
      finally { await handle.close(); }
      await privateMode(target, 0o600);
      return Object.freeze({
        path: target,
        token,
        async release() {
          const current = await readSmallJson(target).catch(() => null);
          if (current?.token === token) await rm(target, { force: true });
        },
      });
    } catch (error) {
      if (error?.code !== "EEXIST") throw error;
      const existing = await readSmallJson(target).catch(() => null);
      if (attempt === 0 && existing && !processExists(existing.pid)) {
        const descriptor = await lstatIfExists(target);
        if (descriptor?.isFile() && !descriptor.isSymbolicLink()) {
          await rm(target, { force: true });
          continue;
        }
      }
      throw taskError("task-busy", `Task is already open by session ${existing?.sessionId || "unknown"}.`);
    }
  }
  throw taskError("task-busy", "Task is already open.");
}

export async function beginTaskSession(opened, { sessionId, now = new Date() }) {
  const manifest = structuredClone(opened.manifest);
  const parentSessionId = manifest.lastSessionId;
  const sessionRoot = await ensurePrivateDirectory(path.join(opened.taskRoot, "sessions", sessionId), opened.taskRoot);
  const interrupted = parentSessionId
    ? await findInterruptedRequest(path.join(opened.taskRoot, "sessions", parentSessionId, "session.jsonl"))
    : null;
  if (interrupted && !manifest.pending.some((entry) => entry.type === "interrupted-request" && entry.sessionId === parentSessionId && entry.sequence === interrupted.sequence)) {
    manifest.pending.push({
      type: "interrupted-request",
      sessionId: parentSessionId,
      sequence: interrupted.sequence,
      requestId: interrupted.id,
      sourceSha256: interrupted.sourceSha256,
      maybeApplied: true,
      at: now.toISOString(),
    });
  }
  await appendSourceAttention(manifest, now);
  const headCommit = manifest.head
    ? manifest.commits.find((commit) => commit.id === manifest.head.commitId)
    : null;
  const operations = [];
  for (const commit of manifest.commits) {
    if (!commit.operation) continue;
    const record = await readTaskOperationRecord(opened.taskRoot, commit.operation);
    operations.push({ commitId: commit.id, artifactId: commit.artifactId, ...taskOperationDescriptor(opened.taskRoot, commit.operation, record) });
  }
  const restoredArtifacts = [];
  for (const artifact of manifest.artifacts) {
    const revision = headCommit?.heads?.[artifact.id];
    if (!revision) continue;
    const revisionPath = resolveManagedFile(opened.taskRoot, revision.path, "revision");
    const bytes = await readRegularBounded(revisionPath, DEFAULT_MAX_TASK_ARTIFACT_BYTES, "Committed revision");
    if (sha256(bytes) !== revision.sha256) throw taskError("revision-corrupt", `Committed revision hash verification failed for ${artifact.id}.`);
    restoredArtifacts.push({
      artifactId: artifact.id,
      name: artifact.name,
      kind: artifact.kind,
      path: revisionPath,
      bytes: revision.bytes,
      sha256: revision.sha256,
      commitId: headCommit.id,
    });
  }
  manifest.lastSessionId = sessionId;
  manifest.updatedAt = now.toISOString();
  await writeTaskManifest(opened.taskRoot, manifest);
  const program = await resumeTaskPpjRevision({ ...opened, manifest });
  return {
    ...opened,
    manifest,
    parentSessionId,
    sessionId,
    sessionRoot,
    ready: {
      protocol: 3,
      type: "session.ready",
      task: summarizeTask(manifest, { detailed: true, taskRoot: opened.taskRoot }),
      program,
      resumedFrom: headCommit ? {
        commitId: headCommit.id,
        summary: headCommit.summary,
        reviewVerdict: headCommit.review.verdict,
        visualReview: headCommit.review.visualReview,
      } : null,
      commit: headCommit ? createCommitDescriptor(manifest, headCommit) : null,
      operations,
      artifacts: restoredArtifacts,
      session: { id: sessionId, parentSessionId },
    },
  };
}

async function appendSourceAttention(manifest, now) {
  for (const artifact of manifest.artifacts) {
    if (!artifact.source) continue;
    const descriptor = await lstatIfExists(artifact.source.path);
    let type;
    let currentSha256;
    if (!descriptor || descriptor.isSymbolicLink() || !descriptor.isFile()) {
      type = "source-unavailable";
    } else if (descriptor.size > DEFAULT_MAX_TASK_ARTIFACT_BYTES) {
      type = "source-unavailable";
    } else {
      const bytes = await readFile(artifact.source.path);
      if (bytes.byteLength > DEFAULT_MAX_TASK_ARTIFACT_BYTES) {
        type = "source-unavailable";
      } else {
        currentSha256 = sha256(bytes);
      }
      if (!type && currentSha256 !== artifact.source.sha256) type = "source-changed";
    }
    if (!type) continue;
    if (manifest.pending.some((entry) => entry.type === type && entry.artifactId === artifact.id && entry.currentSha256 === currentSha256)) continue;
    manifest.pending.push({
      type,
      artifactId: artifact.id,
      summary: type === "source-changed"
        ? `${artifact.name} changed outside OfficeKit after staging`
        : `${artifact.name} is no longer available at its original path`,
      sourceSha256: artifact.source.sha256,
      currentSha256,
      at: now.toISOString(),
    });
  }
}

export async function stageTaskInput(task, sourcePath, options = {}) {
  if (typeof sourcePath !== "string" || sourcePath.trim() === "") {
    throw taskError("invalid-input", "ctx.input requires a non-empty local path.");
  }
  const requested = path.resolve(task.workspaceRoot, sourcePath);
  const sourceStat = await lstatIfExists(requested);
  if (!sourceStat || sourceStat.isSymbolicLink() || !sourceStat.isFile()) {
    throw taskError("unsafe-input", "Task input must be an existing regular non-symlink file.");
  }
  const maximum = positiveInteger(options.maxBytes, DEFAULT_MAX_TASK_ARTIFACT_BYTES, "maxBytes");
  if (sourceStat.size > maximum) throw taskError("input-too-large", `Task input exceeds ${maximum} bytes.`);
  const canonical = await realpath(requested);
  assertOutsideManagedTasks(canonical, task.workspaceRoot);
  const bytes = await readFile(canonical);
  if (bytes.byteLength > maximum) throw taskError("input-too-large", `Task input exceeds ${maximum} bytes.`);
  const digest = sha256(bytes);
  const existing = task.manifest.artifacts.find((artifact) => artifact.source?.path === canonical);
  if (existing) {
    if (existing.source.sha256 !== digest) throw taskError("source-changed", `Task input changed after it was staged: ${canonical}`);
    return artifactDescriptor(existing, task.taskRoot);
  }
  const artifactId = options.artifactId == null ? newArtifactId() : validateArtifactId(options.artifactId);
  if (task.manifest.artifacts.some((artifact) => artifact.id === artifactId)) {
    throw taskError("artifact-exists", `Artifact ID already exists: ${artifactId}`);
  }
  const kind = normalizeKind(options.kind, canonical, options.mime);
  const extension = extensionFor(canonical, kind);
  const relative = toPosix(path.join("inputs", artifactId, `${digest}${extension}`));
  const destination = path.join(task.taskRoot, relative);
  await ensurePrivateDirectory(path.dirname(destination), task.taskRoot);
  await writeImmutable(destination, bytes, 0o400);
  const artifact = {
    id: artifactId,
    name: boundedText(options.name ?? path.basename(canonical), "Artifact name", 255),
    kind,
    mime: options.mime || mimeForKind(kind),
    source: { path: canonical, storedPath: relative, bytes: bytes.byteLength, sha256: digest },
    headRevision: null,
  };
  task.manifest.artifacts.push(artifact);
  task.manifest.updatedAt = new Date().toISOString();
  await writeTaskManifest(task.taskRoot, task.manifest);
  return artifactDescriptor(artifact, task.taskRoot);
}

export async function readTaskPlan(task) {
  const descriptor = task.manifest.plan;
  if (descriptor == null) return null;
  validateStoredPlanDescriptor(descriptor);
  const target = resolveManagedFile(task.taskRoot, descriptor.path, "authoring plan");
  const bytes = await readRegularBounded(target, MAX_AUTHORING_PLAN_BYTES, "Authoring plan");
  if (bytes.byteLength !== descriptor.bytes || sha256(bytes) !== descriptor.sha256) {
    throw taskError("authoring-plan-corrupt", "Authoring plan bytes do not match the task manifest.");
  }
  let value;
  try { value = JSON.parse(bytes.toString("utf8")); }
  catch (error) { throw taskError("authoring-plan-corrupt", `Authoring plan is not valid JSON: ${boundedError(error)}`); }
  const normalized = normalizePresentationAuthoringPlan(value);
  if (normalized.sha256 !== descriptor.sha256 || normalized.plan.mode !== descriptor.mode ||
      normalized.pageCount !== descriptor.pageCount || normalized.plan.recipe !== descriptor.recipe ||
      descriptor.deliveryMode != null && normalized.deliveryMode !== descriptor.deliveryMode ||
      descriptor.motionPolicy != null && normalized.motionPolicy !== descriptor.motionPolicy ||
      descriptor.motionPageCount != null && normalized.motionPageCount !== descriptor.motionPageCount ||
      descriptor.designGrammarSha256 != null && normalized.designGrammarSha256 !== descriptor.designGrammarSha256 ||
      descriptor.strategyStatus != null && normalized.strategyStatus !== descriptor.strategyStatus ||
      descriptor.primaryJob != null && normalized.primaryJob !== descriptor.primaryJob ||
      descriptor.primaryScenario != null && normalized.primaryScenario !== descriptor.primaryScenario ||
      descriptor.directionName != null && normalized.directionName !== descriptor.directionName ||
      descriptor.mediumFit != null && normalized.mediumFit !== descriptor.mediumFit) {
    throw taskError("authoring-plan-corrupt", "Authoring plan content does not match its manifest descriptor.");
  }
  validatePlanArtifactBindings(normalized.plan, task.manifest);
  return structuredClone(normalized.plan);
}

export async function writeTaskPlan(task, value, { expectedSha256, now = new Date() } = {}) {
  const current = task.manifest.plan;
  const normalized = normalizePresentationAuthoringPlan(value, {
    allowLegacy: current != null && (current.strategyStatus == null || current.strategyStatus === "legacy"),
  });
  validatePlanArtifactBindings(normalized.plan, task.manifest);
  if (current != null) {
    if (expectedSha256 !== current.sha256) {
      throw taskError("stale-authoring-plan", "Updating an authoring plan requires its exact current SHA-256.", {
        expectedSha256: current.sha256,
      });
    }
    if (normalized.sha256 === current.sha256) {
      await readTaskPlan(task);
      return Object.freeze({ ...planDescriptorForManifest(task.manifest), unchanged: true });
    }
  } else if (expectedSha256 != null) {
    throw taskError("stale-authoring-plan", "The task has no authoring plan to match expectedSha256.");
  }
  const relative = toPosix(path.join("plans", `${normalized.sha256}.json`));
  const target = path.join(task.taskRoot, relative);
  await ensurePrivateDirectory(path.dirname(target), task.taskRoot);
  await writeImmutable(target, normalized.bytes, 0o400, { allowIdentical: true });
  task.manifest.plan = {
    ...authoringPlanDescriptor(normalized, { path: relative }),
  };
  delete task.manifest.plan.state;
  if (normalized.plan.nextAction != null) task.manifest.next = normalized.plan.nextAction;
  task.manifest.updatedAt = now.toISOString();
  await writeTaskManifest(task.taskRoot, task.manifest);
  return Object.freeze({ ...planDescriptorForManifest(task.manifest), unchanged: false });
}

export async function recordTaskPpjRevision(task, workspace, {
  stage,
  receipt,
  candidate = null,
  review = null,
  now = new Date(),
} = {}) {
  if (!PPJ_RECEIPT_STAGES.has(stage)) throw taskError("invalid-ppj-stage", "PPJ task stage is invalid.");
  if (task.manifest.plan != null) {
    throw taskError(
      "unsupported-task-schema",
      "A legacy ctx.plan presentation task cannot be migrated into a PPJ task. Start a new PPJ task and keep the legacy task read-only.",
    );
  }
  if (!workspace || !receipt || !(receipt.programJson instanceof Uint8Array)) {
    throw taskError("invalid-ppj-receipt", "PPJ task recording requires a validated native receipt and loaded workspace.");
  }
  const programBytes = Buffer.from(receipt.programJson);
  if (programBytes.byteLength === 0 || programBytes.byteLength > DEFAULT_MAX_TASK_PPJ_BYTES ||
      !isSha(receipt.programSha256) || sha256(programBytes) !== receipt.programSha256) {
    throw taskError("invalid-ppj-receipt", "PPJ task receipt does not bind its canonical program bytes.");
  }
  let program;
  try { program = JSON.parse(programBytes.toString("utf8")); }
  catch { throw taskError("invalid-ppj-receipt", "PPJ task receipt is not valid JSON."); }
  if (program?.schema !== PPJ_PROGRAM_SCHEMA || !Array.isArray(program.pages)) {
    throw taskError("invalid-ppj-receipt", "PPJ task receipt does not contain an office-kit/ppj/v1 program.");
  }

  const revisionRootRelative = toPosix(path.join("programs", receipt.programSha256));
  await ensurePrivateSubdirectory(task.taskRoot, revisionRootRelative);
  const programRelative = `${revisionRootRelative}/program.ppj`;
  await writeImmutable(path.join(task.taskRoot, programRelative), programBytes, 0o400, { allowIdentical: true });
  const resources = await storePpjResources(task, workspace, program, revisionRootRelative);

  let nodeMap = null;
  if (receipt.nodeMapJson instanceof Uint8Array && receipt.nodeMapJson.byteLength > 0) {
    const bytes = Buffer.from(receipt.nodeMapJson);
    if (bytes.byteLength > DEFAULT_MAX_TASK_PPJ_BYTES) throw taskError("invalid-ppj-receipt", "PPJ node map exceeds its task budget.");
    const digest = sha256(bytes);
    const relative = `${revisionRootRelative}/node-map.json`;
    await writeImmutable(path.join(task.taskRoot, relative), bytes, 0o400, { allowIdentical: true });
    nodeMap = { path: relative, bytes: bytes.byteLength, sha256: digest };
  }

  const candidateDescriptor = candidate == null
    ? null
    : await storePpjCandidate(task, receipt, candidate);
  const reviewDescriptor = review == null
    ? null
    : await storePpjReview(task, receipt, candidateDescriptor, review);
  if (stage === "reviewed" && (candidateDescriptor == null || reviewDescriptor == null)) {
    throw taskError("invalid-ppj-review", "A reviewed PPJ revision requires its exact candidate and review report.");
  }

  const receiptValue = {
    schema: "office-kit/ppj-task-receipt/v1",
    stage,
    programSha256: receipt.programSha256,
    sourceBound: Boolean(receipt.sourceBound),
    restoredEmbeddedProgram: Boolean(receipt.restoredEmbeddedProgram),
    sourceSha256: isSha(receipt.sourceSha256) ? receipt.sourceSha256 : null,
    outputSha256: isSha(receipt.outputSha256) ? receipt.outputSha256 : null,
    expandedElementCount: Number(receipt.expandedElementCount ?? 0),
    changedParts: [...(receipt.changedParts ?? [])],
    changedNodeIds: [...(receipt.changedNodeIds ?? [])],
    nodeMapSha256: nodeMap?.sha256 ?? null,
    diagnostics: (receipt.diagnostics ?? []).map((diagnostic) => ({
      severity: diagnostic.severity,
      code: diagnostic.code,
      message: diagnostic.message,
      sourcePath: diagnostic.sourcePath,
      sourceIdentity: diagnostic.sourceIdentity,
    })),
  };
  const receiptBytes = Buffer.from(`${JSON.stringify(receiptValue, null, 2)}\n`);
  const receiptSha256 = sha256(receiptBytes);
  const receiptRelative = toPosix(path.join("evidence", "ppj", receipt.programSha256, `${stage}-${receiptSha256}.json`));
  await ensurePrivateSubdirectory(task.taskRoot, path.dirname(receiptRelative));
  await writeImmutable(path.join(task.taskRoot, receiptRelative), receiptBytes, 0o400, { allowIdentical: true });
  const receiptDescriptor = { stage, path: receiptRelative, bytes: receiptBytes.byteLength, sha256: receiptSha256, recordedAt: now.toISOString() };

  const ppj = task.manifest.ppj ?? { schema: PPJ_TASK_SCHEMA, head: null, reviewed: null, revisions: [] };
  if (ppj.schema !== PPJ_TASK_SCHEMA || !Array.isArray(ppj.revisions)) throw taskError("invalid-task", "Task PPJ state is invalid.");
  let revision = ppj.revisions.find((item) => item.sha256 === receipt.programSha256);
  const identity = {
    mode: receipt.sourceBound ? "source-bound" : receipt.restoredEmbeddedProgram ? "embedded-authored" : "authored",
    sourceSha256: isSha(receipt.sourceSha256) ? receipt.sourceSha256 : null,
  };
  if (!revision) {
    if (ppj.revisions.length >= MAX_PPJ_REVISIONS) throw taskError("ppj-revision-budget", `Task exceeds ${MAX_PPJ_REVISIONS} PPJ revisions.`);
    revision = {
      sha256: receipt.programSha256,
      bytes: programBytes.byteLength,
      path: programRelative,
      identity,
      resources,
      nodeMap,
      receipts: [],
      candidate: null,
      review: null,
      status: "valid",
      updatedAt: now.toISOString(),
    };
    ppj.revisions.push(revision);
  } else {
    if (revision.path !== programRelative || revision.bytes !== programBytes.byteLength ||
        revision.identity?.mode !== identity.mode || revision.identity?.sourceSha256 !== identity.sourceSha256 ||
        JSON.stringify(revision.resources) !== JSON.stringify(resources)) {
      throw taskError("ppj-revision-collision", "Existing PPJ task revision has incompatible immutable bindings.");
    }
    if (nodeMap && revision.nodeMap && revision.nodeMap.sha256 !== nodeMap.sha256) {
      throw taskError("ppj-revision-collision", "Existing PPJ task revision has a different node map.");
    }
    revision.nodeMap ??= nodeMap;
  }
  const previousReceipt = revision.receipts.find((item) => item.stage === stage);
  if (previousReceipt && previousReceipt.sha256 !== receiptDescriptor.sha256) {
    throw taskError("ppj-revision-collision", `PPJ ${stage} receipt changed for the same immutable program revision.`);
  }
  if (!previousReceipt) revision.receipts.push(receiptDescriptor);
  if (candidateDescriptor && revision.candidate && revision.candidate.sha256 !== candidateDescriptor.sha256) {
    throw taskError("ppj-revision-collision", "PPJ compiler output changed for the same immutable program revision.");
  }
  if (candidateDescriptor && candidateDescriptor.outputPath == null && revision.candidate?.outputPath != null) {
    candidateDescriptor.outputPath = revision.candidate.outputPath;
  }
  if (candidateDescriptor) revision.candidate = candidateDescriptor;
  if (reviewDescriptor) revision.review = reviewDescriptor;
  revision.status = revision.review
    ? revision.review.verdict === "failed" ? "review-failed" : "reviewed"
    : revision.candidate ? "candidate" : "valid";
  revision.updatedAt = now.toISOString();
  ppj.head = revision.sha256;
  if (revision.status === "reviewed") ppj.reviewed = revision.sha256;
  task.manifest.ppj = ppj;
  task.manifest.updatedAt = now.toISOString();
  if (revision.status === "review-failed") {
    task.manifest.pending.push({
      type: "ppj-review-failed",
      summary: `PPJ ${revision.sha256.slice(0, 12)} did not pass review`,
      programSha256: revision.sha256,
      at: now.toISOString(),
    });
  } else {
    task.manifest.pending = task.manifest.pending.filter((item) =>
      item.type !== "ppj-review-failed" || item.programSha256 !== revision.sha256);
  }
  await writeTaskManifest(task.taskRoot, task.manifest);
  return Object.freeze(ppjRevisionDescriptor(task.manifest, task.taskRoot, { detailed: true }));
}

export async function resumeTaskPpjRevision(task) {
  if (task.manifest.ppj == null) {
    if (task.manifest.plan != null) return Object.freeze({
      status: "unsupported",
      code: "unsupported-task-schema",
      schema: task.manifest.plan.schema,
      message: "Legacy ctx.plan tasks remain listable but are not migrated into PPJ 2.0.",
    });
    return null;
  }
  const ppj = task.manifest.ppj;
  const candidates = [...new Set([ppj.head, ppj.reviewed].filter(Boolean))];
  const failures = [];
  for (const sha of candidates) {
    const revision = ppj.revisions.find((item) => item.sha256 === sha);
    if (!revision) continue;
    try {
      await verifyPpjTaskRevision(task.taskRoot, revision);
      return Object.freeze({
        ...ppjRevisionDescriptor(task.manifest, task.taskRoot, { detailed: true, revisionSha256: sha }),
        resumedFromFallback: sha !== ppj.head,
      });
    } catch (error) {
      failures.push({ sha256: sha, message: boundedError(error) });
    }
  }
  throw taskError("ppj-revision-corrupt", "No valid PPJ task revision is available to resume.", { failures });
}

export async function commitTaskArtifact(task, value, options = {}) {
  const bytes = await readArtifactBytes(value, options.maxBytes);
  const digest = sha256(bytes);
  const editPlan = validateTaskEditPlan(value?.metadata?.editPlan, digest);
  const artifactId = validateArtifactId(options.artifactId);
  let artifact = task.manifest.artifacts.find((entry) => entry.id === artifactId);
  const review = validateReview(options.review);
  const summary = boundedText(options.summary, "Commit summary", 1_024);
  const next = options.next == null ? null : boundedText(options.next, "Next action", 1_024);
  const constraints = options.constraints == null
    ? null
    : normalizeConstraints(options.constraints);
  const kind = normalizeKind(options.kind ?? review.artifactKind, options.name, options.mime);
  if (task.manifest.plan && kind === "presentation" && review.design?.planSha256 !== task.manifest.plan.sha256) {
    throw taskError("stale-authoring-plan-review", "Presentation review must be bound to the active authoring plan SHA-256.", {
      expectedPlanSha256: task.manifest.plan.sha256,
      reviewPlanSha256: review.design?.planSha256 ?? null,
    });
  }
  if (!artifact) {
    artifact = {
      id: artifactId,
      name: boundedText(options.name ?? `${artifactId}${extensionFor("", kind)}`, "Artifact name", 255),
      kind,
      mime: options.mime || mimeForKind(kind),
      source: null,
      headRevision: null,
    };
    task.manifest.artifacts.push(artifact);
  } else if (artifact.kind !== kind) {
    throw taskError("artifact-kind-mismatch", `Artifact ${artifactId} is ${artifact.kind}, not ${kind}.`);
  }
  if (review.delivery.sha256 !== digest) {
    const pending = await storePendingCandidate(task, artifact, bytes, digest, options, review, "stale-review");
    throw taskError("stale-review", "Review delivery SHA-256 does not match the candidate.", { pending });
  }
  if (review.verdict === "failed") {
    const pending = await storePendingCandidate(task, artifact, bytes, digest, options, review, "review-failed");
    throw taskError("review-failed", "A failed review cannot advance task HEAD.", { pending });
  }
  if (!new Set(["passed", "passed-with-limitations"]).has(review.verdict)) {
    throw taskError("invalid-review", "Review verdict must be passed, passed-with-limitations, or failed.");
  }
  const commitId = `c${String(task.manifest.commits.length + 1).padStart(4, "0")}`;
  const extension = extensionFor(artifact.name, artifact.kind);
  const revisionRelative = toPosix(path.join("revisions", artifact.id, `${digest}${extension}`));
  const revisionPath = path.join(task.taskRoot, revisionRelative);
  await ensurePrivateDirectory(path.dirname(revisionPath), task.taskRoot);
  await writeImmutable(revisionPath, bytes, 0o400, { allowIdentical: true });
  const report = await writeReviewEvidence(task, `${commitId}-${artifact.id}`, options.review);
  const revisionReview = reviewSummary(review, report);
  const operation = editPlan
    ? await writeTaskOperationRecord(task, commitId, artifact.id, editPlan)
    : null;
  artifact.headRevision = { sha256: digest, bytes: bytes.byteLength, path: revisionRelative, commitId, review: revisionReview };
  const heads = Object.fromEntries(task.manifest.artifacts
    .filter((entry) => entry.headRevision)
    .map((entry) => [entry.id, structuredClone(entry.headRevision)]));
  const committedAt = new Date().toISOString();
  const commit = {
    id: commitId,
    artifactId: artifact.id,
    revisionSha256: digest,
    summary,
    next,
    committedAt,
    heads,
    review: revisionReview,
    plan: task.manifest.plan ? structuredClone(task.manifest.plan) : null,
    ...(operation ? { operation } : {}),
  };
  task.manifest.commits.push(commit);
  task.manifest.head = { commitId, artifactId: artifact.id, revisionSha256: digest, committedAt };
  task.manifest.next = next;
  if (constraints) task.manifest.constraints = constraints;
  task.manifest.updatedAt = committedAt;
  task.manifest.pending = task.manifest.pending.filter((entry) => {
    if (entry.type === "interrupted-request") return false;
    return entry.artifactId !== artifact.id || !RESOLVED_BY_COMMIT.has(entry.type);
  });
  await writeTaskManifest(task.taskRoot, task.manifest);
  return Object.freeze(createCommitDescriptor(task.manifest, commit));
}

export async function resolveCommittedArtifact(task, descriptor, requestedArtifactId = descriptor?.artifactId) {
  if (descriptor?.type !== "officekit.task-commit" || descriptor.taskId !== task.manifest.id) {
    throw taskError("unreviewed-artifact", "ctx.publish accepts only a commit from the current task.");
  }
  const head = task.manifest.head;
  if (!head || head.commitId !== descriptor.commitId) {
    throw taskError("stale-commit", "Only the current reviewed task commit can be published.");
  }
  validateArtifactId(requestedArtifactId);
  const commit = task.manifest.commits.find((entry) => entry.id === descriptor.commitId);
  if (task.manifest.plan && commit?.plan?.sha256 !== task.manifest.plan.sha256) {
    throw taskError("unreviewed-authoring-plan", "The active authoring plan is newer than the current reviewed artifact commit.");
  }
  if (commit?.heads?.[descriptor.artifactId]?.sha256 !== descriptor.revisionSha256) {
    throw taskError("invalid-commit", "Commit descriptor does not match the reviewed task commit.");
  }
  const artifact = task.manifest.artifacts.find((entry) => entry.id === requestedArtifactId);
  const revision = commit?.heads?.[artifact?.id];
  if (!commit || !artifact || !revision) {
    throw taskError("invalid-commit", "Committed artifact metadata is incomplete or inconsistent.");
  }
  const revisionPath = resolveManagedFile(task.taskRoot, revision.path, "revision");
  const bytes = await readRegularBounded(revisionPath, DEFAULT_MAX_TASK_ARTIFACT_BYTES, "Committed revision");
  if (sha256(bytes) !== revision.sha256) throw taskError("revision-corrupt", "Committed revision hash verification failed.");
  return { artifact, commit, revision, review: revision.review ?? commit.review, bytes };
}

export async function recordTaskPublication(task, commitDescriptor, publication, artifactId = commitDescriptor?.artifactId) {
  const committed = await resolveCommittedArtifact(task, commitDescriptor, artifactId);
  task.manifest.publications.push({
    commitId: committed.commit.id,
    artifactId: committed.artifact.id,
    path: publication.path,
    bytes: publication.bytes,
    sha256: publication.sha256,
    publishedAt: new Date().toISOString(),
  });
  task.manifest.updatedAt = new Date().toISOString();
  await writeTaskManifest(task.taskRoot, task.manifest);
}

export function summarizeTask(manifest, { detailed = false, taskRoot } = {}) {
  validateTaskManifest(manifest, manifest.id);
  const state = deriveTaskState(manifest);
  const headCommit = manifest.head
    ? manifest.commits.find((commit) => commit.id === manifest.head.commitId)
    : null;
  const base = {
    id: manifest.id,
    goal: manifest.goal,
    head: headCommit ? {
      id: headCommit.id,
      summary: headCommit.summary,
      reviewVerdict: headCommit.review.verdict,
      visualReview: headCommit.review.visualReview,
      committedAt: headCommit.committedAt,
    } : null,
    plan: planDescriptorForManifest(manifest),
    program: manifest.ppj != null
      ? ppjRevisionDescriptor(manifest, taskRoot, { detailed })
      : manifest.plan != null ? {
        status: "unsupported",
        code: "unsupported-task-schema",
        schema: manifest.plan.schema,
        message: "Legacy ctx.plan tasks remain listable but are not migrated into PPJ 2.0.",
      } : null,
    state,
    updatedAt: manifest.updatedAt,
  };
  if (!detailed) return base;
  return {
    ...base,
    createdAt: manifest.createdAt,
    inputs: manifest.artifacts.filter((artifact) => artifact.source).map((artifact) => ({
      artifactId: artifact.id,
      name: artifact.name,
      kind: artifact.kind,
      path: artifact.source.path,
      bytes: artifact.source.bytes,
      sha256: artifact.source.sha256,
    })),
    artifacts: manifest.artifacts.map((artifact) => ({
      id: artifact.id,
      name: artifact.name,
      kind: artifact.kind,
      headRevision: artifact.headRevision ? structuredClone(artifact.headRevision) : null,
    })),
    pending: structuredClone(manifest.pending),
    next: manifest.next,
    constraints: structuredClone(manifest.constraints),
    commit: headCommit ? createCommitDescriptor(manifest, headCommit) : null,
    publication: manifest.publications.at(-1) ? structuredClone(manifest.publications.at(-1)) : null,
    storageBytes: taskStorageBytes(manifest),
  };
}

function createCommitDescriptor(manifest, commit) {
  return {
    type: "officekit.task-commit",
    taskId: manifest.id,
    commitId: commit.id,
    artifactId: commit.artifactId,
    revisionSha256: commit.revisionSha256,
    reviewVerdict: commit.review.verdict,
    visualReview: commit.review.visualReview,
    operation: commit.operation ? structuredClone(commit.operation) : null,
    plan: commit.plan ? Object.freeze({ ...commit.plan, state: "reviewed" }) : null,
    artifacts: Object.entries(commit.heads).map(([artifactId, revision]) => ({ artifactId, sha256: revision.sha256 })),
  };
}

function planDescriptorForManifest(manifest) {
  if (manifest.plan == null) return null;
  const headCommit = manifest.head
    ? manifest.commits.find((commit) => commit.id === manifest.head.commitId)
    : null;
  return Object.freeze({
    ...structuredClone(manifest.plan),
    strategyStatus: manifest.plan.strategyStatus ?? "legacy",
    state: headCommit?.plan?.sha256 === manifest.plan.sha256 ? "reviewed" : "working",
  });
}

function validateStoredPlanDescriptor(descriptor) {
  if (!descriptor || descriptor.schema !== PRESENTATION_AUTHORING_PLAN_SCHEMA ||
      typeof descriptor.mode !== "string" || descriptor.mode.length === 0 || descriptor.mode.length > 64 ||
      !Number.isSafeInteger(descriptor.pageCount) || descriptor.pageCount <= 0 || descriptor.pageCount > 64 ||
      typeof descriptor.recipe !== "string" || descriptor.recipe.length === 0 || descriptor.recipe.length > 160 ||
      descriptor.deliveryMode != null && !new Set(["live", "reader", "hybrid"]).has(descriptor.deliveryMode) ||
      descriptor.motionPolicy != null && !new Set(["adaptive", "none", "explicit"]).has(descriptor.motionPolicy) ||
      descriptor.motionPageCount != null && (!Number.isSafeInteger(descriptor.motionPageCount) || descriptor.motionPageCount < 0 || descriptor.motionPageCount > descriptor.pageCount) ||
      descriptor.designGrammarSha256 != null && !isSha(descriptor.designGrammarSha256) ||
      descriptor.strategyStatus != null && !PRESENTATION_STRATEGY_STATUS_SET.has(descriptor.strategyStatus) ||
      descriptor.primaryJob != null && !PRESENTATION_COMMUNICATION_JOB_SET.has(descriptor.primaryJob) ||
      descriptor.primaryScenario != null && !PRESENTATION_SCENARIO_SET.has(descriptor.primaryScenario) ||
      descriptor.directionName != null && (typeof descriptor.directionName !== "string" || descriptor.directionName.trim() === "" || descriptor.directionName.length > 160) ||
      descriptor.mediumFit != null && !PRESENTATION_MEDIUM_FIT_SET.has(descriptor.mediumFit) ||
      descriptor.strategyStatus === "current" &&
        (descriptor.primaryJob == null || descriptor.primaryScenario == null || descriptor.directionName == null || descriptor.mediumFit == null) ||
      !isSha(descriptor.sha256) || !Number.isSafeInteger(descriptor.bytes) || descriptor.bytes <= 0 || descriptor.bytes > MAX_AUTHORING_PLAN_BYTES ||
      typeof descriptor.path !== "string") {
    throw taskError("invalid-task", "Task authoring-plan descriptor is invalid.");
  }
}

function validatePlanArtifactBindings(plan, manifest) {
  const available = new Map();
  for (const artifact of manifest.artifacts) {
    const hashes = new Set();
    if (artifact.source?.sha256) hashes.add(artifact.source.sha256);
    if (artifact.headRevision?.sha256) hashes.add(artifact.headRevision.sha256);
    for (const commit of manifest.commits) {
      const revision = commit.heads?.[artifact.id];
      if (revision?.sha256) hashes.add(revision.sha256);
    }
    available.set(artifact.id, hashes);
  }
  for (const ref of plan.artifactRefs ?? []) {
    if (!available.get(ref.artifactId)?.has(ref.sha256)) {
      throw taskError("unbound-authoring-plan-reference", `Authoring plan reference ${ref.artifactId}@${ref.sha256} is not a task artifact revision.`);
    }
  }
}

async function validateTaskPlanFiles(manifest, taskRoot) {
  const descriptors = new Map();
  if (manifest.plan) descriptors.set(manifest.plan.sha256, manifest.plan);
  for (const commit of manifest.commits) {
    if (commit.plan) descriptors.set(commit.plan.sha256, commit.plan);
  }
  for (const descriptor of descriptors.values()) {
    const target = resolveManagedFile(taskRoot, descriptor.path, "authoring plan");
    const bytes = await readRegularBounded(target, MAX_AUTHORING_PLAN_BYTES, "Authoring plan");
    if (bytes.byteLength !== descriptor.bytes || sha256(bytes) !== descriptor.sha256) {
      throw taskError("authoring-plan-corrupt", "Authoring plan hash verification failed.");
    }
    let value;
    try { value = JSON.parse(bytes.toString("utf8")); }
    catch (error) { throw taskError("authoring-plan-corrupt", `Authoring plan is not valid JSON: ${boundedError(error)}`); }
    const normalized = normalizePresentationAuthoringPlan(value);
    if (normalized.sha256 !== descriptor.sha256) throw taskError("authoring-plan-corrupt", "Authoring plan canonical hash verification failed.");
    validatePlanArtifactBindings(normalized.plan, manifest);
  }
}

async function ensureTaskStore(workspaceRoot) {
  const officeKit = await ensurePrivateDirectory(path.join(workspaceRoot, ".office-kit"), workspaceRoot);
  const tasksRoot = await ensurePrivateDirectory(path.join(officeKit, "tasks"), officeKit);
  const ignorePath = path.join(tasksRoot, ".gitignore");
  const existing = await lstatIfExists(ignorePath);
  if (existing == null) await writeFile(ignorePath, TASK_IGNORE, { encoding: "utf8", mode: 0o600 });
  else if (existing.isSymbolicLink() || !existing.isFile()) throw taskError("unsafe-task-store", "Task ignore marker must be a regular file.");
  return tasksRoot;
}

async function resolveTaskRoot(workspaceRoot, taskId) {
  validateTaskId(taskId);
  const tasksRoot = path.join(workspaceRoot, TASK_DIRECTORY);
  const tasksStat = await lstatIfExists(tasksRoot);
  if (!tasksStat || tasksStat.isSymbolicLink() || !tasksStat.isDirectory()) {
    throw taskError("task-not-found", `OfficeKit task does not exist: ${taskId}`);
  }
  const canonicalTasks = await realpath(tasksRoot);
  const candidate = path.join(canonicalTasks, taskId);
  const descriptor = await lstatIfExists(candidate);
  if (!descriptor || descriptor.isSymbolicLink() || !descriptor.isDirectory()) {
    throw taskError("task-not-found", `OfficeKit task does not exist: ${taskId}`);
  }
  const canonical = await realpath(candidate);
  assertContained(canonical, canonicalTasks, "task");
  return canonical;
}

async function readTaskManifest(taskRoot, expectedId) {
  const manifestPath = path.join(taskRoot, "task.json");
  const bytes = await readRegularBounded(manifestPath, DEFAULT_MAX_TASK_MANIFEST_BYTES, "Task manifest");
  let manifest;
  try { manifest = JSON.parse(bytes.toString("utf8")); }
  catch (error) { throw taskError("invalid-task", `Task manifest is not valid JSON: ${error.message}`); }
  const normalized = normalizeTaskManifestForRead(manifest, expectedId);
  validateTaskPaths(normalized, taskRoot);
  await validateTaskOperationFiles(normalized, taskRoot);
  await validateTaskPlanFiles(normalized, taskRoot);
  return normalized;
}

async function writeTaskManifest(taskRoot, manifest) {
  validateTaskManifest(manifest, manifest.id);
  const serialized = `${JSON.stringify(manifest, null, 2)}\n`;
  if (Buffer.byteLength(serialized) > DEFAULT_MAX_TASK_MANIFEST_BYTES) throw taskError("task-too-large", "Task manifest exceeds its safety limit.");
  await atomicWrite(path.join(taskRoot, "task.json"), Buffer.from(serialized), taskRoot, { replace: true });
}

function validateTaskManifest(manifest, expectedId) {
  if (!manifest || typeof manifest !== "object" || Array.isArray(manifest) || manifest.schemaVersion !== TASK_SCHEMA_VERSION || manifest.id !== expectedId || !TASK_ID_PATTERN.test(manifest.id)) {
    throw taskError("invalid-task", "Task manifest schema or ID is invalid.");
  }
  boundedText(manifest.goal, "Task goal", 1_024);
  for (const key of ["artifacts", "commits", "pending", "publications", "constraints"]) {
    if (!Array.isArray(manifest[key])) throw taskError("invalid-task", `Task manifest ${key} must be an array.`);
  }
  normalizeConstraints(manifest.constraints);
  if (typeof manifest.createdAt !== "string" || typeof manifest.updatedAt !== "string") throw taskError("invalid-task", "Task timestamps are invalid.");
  if (manifest.plan != null) validateStoredPlanDescriptor(manifest.plan);
  if (manifest.ppj != null) validatePpjTaskDescriptor(manifest.ppj);
  const artifactIds = new Set();
  for (const artifact of manifest.artifacts) {
    validateArtifactId(artifact?.id);
    if (artifactIds.has(artifact.id)) throw taskError("invalid-task", "Task artifact IDs must be unique.");
    artifactIds.add(artifact.id);
    if (!ARTIFACT_KINDS.has(artifact.kind) || typeof artifact.name !== "string") throw taskError("invalid-task", "Task artifact metadata is invalid.");
    if (artifact.source) validateManagedRecord(artifact.source, "input");
    if (artifact.headRevision) validateManagedRecord(artifact.headRevision, "revision");
  }
  const commitIds = new Set();
  for (const commit of manifest.commits) {
    if (!/^c\d{4,}$/u.test(commit?.id) || commitIds.has(commit.id) || !artifactIds.has(commit.artifactId)) throw taskError("invalid-task", "Task commit metadata is invalid.");
    commitIds.add(commit.id);
    if (commit.operation) validateTaskOperationSummary(commit.operation);
    if (commit.plan != null) validateStoredPlanDescriptor(commit.plan);
    if (commit.review?.planSha256 != null && commit.review.planSha256 !== commit.plan?.sha256) {
      throw taskError("invalid-task", "Task commit review and authoring-plan bindings do not match.");
    }
  }
  if (manifest.head && (!commitIds.has(manifest.head.commitId) || !artifactIds.has(manifest.head.artifactId) || !isSha(manifest.head.revisionSha256))) {
    throw taskError("invalid-task", "Task HEAD is invalid.");
  }
  return manifest;
}

function normalizeTaskManifestForRead(manifest, expectedId) {
  if (manifest?.schemaVersion === LEGACY_TASK_SCHEMA_VERSION) {
    const normalized = structuredClone(manifest);
    normalized.schemaVersion = TASK_SCHEMA_VERSION;
    normalized.plan = null;
    normalized.ppj = null;
    for (const commit of normalized.commits ?? []) {
      if (!("plan" in commit)) commit.plan = null;
    }
    validateTaskManifest(normalized, expectedId);
    return normalized;
  }
  if (manifest?.schemaVersion === TASK_SCHEMA_VERSION && !("ppj" in manifest)) {
    const normalized = structuredClone(manifest);
    normalized.ppj = null;
    validateTaskManifest(normalized, expectedId);
    return normalized;
  }
  validateTaskManifest(manifest, expectedId);
  return manifest;
}

function validateTaskPaths(manifest, taskRoot) {
  const artifacts = new Map(manifest.artifacts.map((artifact) => [artifact.id, artifact]));
  for (const artifact of manifest.artifacts) {
    if (artifact.source) {
      validateRelativePrefix(artifact.source.storedPath, `inputs/${artifact.id}/`, taskRoot, "input");
      if (typeof artifact.source.path !== "string" || !path.isAbsolute(artifact.source.path)) throw taskError("invalid-task", "Task source path must be absolute.");
    }
    if (artifact.headRevision) validateRelativePrefix(artifact.headRevision.path, `revisions/${artifact.id}/`, taskRoot, "revision");
  }
  for (const commit of manifest.commits) {
    if (!commit.heads || typeof commit.heads !== "object" || Array.isArray(commit.heads)) throw taskError("invalid-task", "Task commit heads are invalid.");
    for (const [artifactId, revision] of Object.entries(commit.heads)) {
      if (!artifacts.has(artifactId)) throw taskError("invalid-task", "Task commit refers to an unknown artifact.");
      validateManagedRecord(revision, "revision");
      validateRelativePrefix(revision.path, `revisions/${artifactId}/`, taskRoot, "revision");
    }
    validateRelativePrefix(commit.review?.evidence?.path, "evidence/reviews/", taskRoot, "review evidence");
    if (commit.operation) validateRelativePrefix(commit.operation.path, "operations/", taskRoot, "operation record");
    if (commit.plan) validateRelativePrefix(commit.plan.path, "plans/", taskRoot, "authoring plan");
  }
  if (manifest.plan) validateRelativePrefix(manifest.plan.path, "plans/", taskRoot, "authoring plan");
  for (const revision of manifest.ppj?.revisions ?? []) {
    const root = `programs/${revision.sha256}`;
    if (revision.path !== `${root}/program.ppj`) throw taskError("invalid-task", "Task PPJ program path is invalid.");
    validateRelativePrefix(revision.path, `${root}/`, taskRoot, "PPJ revision");
    for (const resource of revision.resources) {
      if (resource.path !== `${root}/${resource.uri}`) throw taskError("invalid-task", "Task PPJ resource path does not match its URI.");
      validateRelativePrefix(resource.path, `${root}/`, taskRoot, "PPJ resource");
    }
    if (revision.nodeMap) {
      if (revision.nodeMap.path !== `${root}/node-map.json`) throw taskError("invalid-task", "Task PPJ node-map path is invalid.");
      validateRelativePrefix(revision.nodeMap.path, `${root}/`, taskRoot, "PPJ node map");
    }
    for (const receipt of revision.receipts) validateRelativePrefix(receipt.path, `evidence/ppj/${revision.sha256}/`, taskRoot, "PPJ receipt");
    if (revision.candidate) validateRelativePrefix(revision.candidate.path, `candidates/ppj/${revision.sha256}/`, taskRoot, "PPJ candidate");
    if (revision.review) validateRelativePrefix(revision.review.evidence.path, `evidence/ppj/${revision.sha256}/`, taskRoot, "PPJ review");
  }
  for (const pending of manifest.pending) {
    if (pending.path != null) validateRelativePrefix(pending.path, "candidates/", taskRoot, "candidate");
    if (pending.review?.evidence?.path != null) validateRelativePrefix(pending.review.evidence.path, "evidence/reviews/", taskRoot, "review evidence");
  }
}

function validateRelativePrefix(value, prefix, taskRoot, label) {
  if (typeof value !== "string" || value === "" || path.isAbsolute(value) || !value.startsWith(prefix) || value.includes("\\")) {
    throw taskError("invalid-task", `Task ${label} path is invalid.`);
  }
  resolveManagedFile(taskRoot, value, label);
}

function validateManagedRecord(record, label) {
  if (!record || !isSha(record.sha256) || !Number.isSafeInteger(record.bytes) || record.bytes < 0 || typeof record.path !== "string" && typeof record.storedPath !== "string") {
    throw taskError("invalid-task", `Task ${label} record is invalid.`);
  }
}

function validatePpjTaskDescriptor(ppj) {
  if (!ppj || typeof ppj !== "object" || Array.isArray(ppj) || ppj.schema !== PPJ_TASK_SCHEMA ||
      !Array.isArray(ppj.revisions) || ppj.revisions.length === 0 || ppj.revisions.length > MAX_PPJ_REVISIONS ||
      !isSha(ppj.head) || ppj.reviewed != null && !isSha(ppj.reviewed)) {
    throw taskError("invalid-task", "Task PPJ descriptor is invalid.");
  }
  const revisions = new Map();
  for (const revision of ppj.revisions) {
    if (!revision || typeof revision !== "object" || Array.isArray(revision) || !isSha(revision.sha256) ||
        revisions.has(revision.sha256) || !Number.isSafeInteger(revision.bytes) || revision.bytes <= 0 ||
        revision.bytes > DEFAULT_MAX_TASK_PPJ_BYTES || typeof revision.path !== "string" ||
        !PPJ_REVISION_STATUSES.has(revision.status) || typeof revision.updatedAt !== "string" ||
        !Array.isArray(revision.resources) || !Array.isArray(revision.receipts) || revision.receipts.length === 0 ||
        revision.receipts.length > PPJ_RECEIPT_STAGES.size) {
      throw taskError("invalid-task", "Task PPJ revision metadata is invalid.");
    }
    const identity = revision.identity;
    if (!identity || !new Set(["authored", "embedded-authored", "source-bound"]).has(identity.mode) ||
        (identity.mode === "source-bound" ? !isSha(identity.sourceSha256) : identity.sourceSha256 != null)) {
      throw taskError("invalid-task", "Task PPJ revision identity is invalid.");
    }
    const resourcePaths = new Set();
    const resourceUris = new Set();
    let sourceCount = 0;
    for (const resource of revision.resources) {
      if (!resource || !new Set(["asset", "source"]).has(resource.kind) ||
          resource.kind === "asset" && (typeof resource.id !== "string" || resource.id === "") ||
          resource.kind === "source" && resource.id != null || typeof resource.mimeType !== "string" ||
          !isSha(resource.sha256) || !Number.isSafeInteger(resource.bytes) || resource.bytes <= 0 ||
          resource.bytes > DEFAULT_MAX_TASK_ARTIFACT_BYTES || typeof resource.path !== "string" ||
          safePpjTaskUri(resource.uri) !== resource.uri || resourcePaths.has(resource.path) || resourceUris.has(resource.uri)) {
        throw taskError("invalid-task", "Task PPJ resource metadata is invalid.");
      }
      resourcePaths.add(resource.path);
      resourceUris.add(resource.uri);
      if (resource.kind === "source") sourceCount += 1;
    }
    if (sourceCount !== (identity.mode === "source-bound" ? 1 : 0)) {
      throw taskError("invalid-task", "Task PPJ source binding does not match its revision identity.");
    }
    if (revision.nodeMap) validateManagedRecord(revision.nodeMap, "PPJ node map");
    const receiptStages = new Set();
    for (const receipt of revision.receipts) {
      validateManagedRecord(receipt, "PPJ receipt");
      if (!PPJ_RECEIPT_STAGES.has(receipt.stage) || receiptStages.has(receipt.stage) || typeof receipt.recordedAt !== "string") {
        throw taskError("invalid-task", "Task PPJ receipt metadata is invalid.");
      }
      receiptStages.add(receipt.stage);
    }
    if (revision.candidate) {
      validateManagedRecord(revision.candidate, "PPJ candidate");
      if (revision.candidate.outputPath != null && (typeof revision.candidate.outputPath !== "string" || !path.isAbsolute(revision.candidate.outputPath))) {
        throw taskError("invalid-task", "Task PPJ candidate output path is invalid.");
      }
    }
    if (revision.review) {
      if (!new Set(["passed", "passed-with-limitations", "failed"]).has(revision.review.verdict) ||
          !VISUAL_REVIEW_STATUSES.has(revision.review.visualReview) ||
          !new Set(["structural", "keynote", "powerpoint"]).has(revision.review.playbackEvidence) ||
          revision.review.candidateSha256 !== revision.candidate?.sha256) {
        throw taskError("invalid-task", "Task PPJ review metadata is invalid.");
      }
      validateManagedRecord(revision.review.evidence, "PPJ review evidence");
    }
    if (new Set(["candidate", "reviewed", "review-failed"]).has(revision.status) !== Boolean(revision.candidate) ||
        new Set(["reviewed", "review-failed"]).has(revision.status) !== Boolean(revision.review) ||
        revision.status === "reviewed" && revision.review?.verdict === "failed" ||
        revision.status === "review-failed" && revision.review?.verdict !== "failed") {
      throw taskError("invalid-task", "Task PPJ revision state is inconsistent.");
    }
    revisions.set(revision.sha256, revision);
  }
  if (!revisions.has(ppj.head) || ppj.reviewed != null && revisions.get(ppj.reviewed)?.status !== "reviewed") {
    throw taskError("invalid-task", "Task PPJ head or reviewed revision is invalid.");
  }
}

async function storePendingCandidate(task, artifact, bytes, digest, options, review, reason) {
  const extension = extensionFor(artifact.name, artifact.kind);
  const nonce = randomUUID().replaceAll("-", "").slice(0, 12);
  const relative = toPosix(path.join("candidates", artifact.id, `${Date.now()}-${nonce}-${digest}${extension}`));
  const destination = path.join(task.taskRoot, relative);
  await ensurePrivateDirectory(path.dirname(destination), task.taskRoot);
  await writeImmutable(destination, bytes, 0o400);
  const report = await writeReviewEvidence(task, `candidate-${artifact.id}-${digest.slice(0, 12)}-${nonce}`, options.review);
  const pending = {
    type: reason,
    artifactId: artifact.id,
    summary: typeof options.summary === "string" ? options.summary.slice(0, 1_024) : "Candidate did not pass review",
    path: relative,
    bytes: bytes.byteLength,
    sha256: digest,
    review: reviewSummary(review, report),
    at: new Date().toISOString(),
  };
  task.manifest.pending.push(pending);
  task.manifest.updatedAt = pending.at;
  await writeTaskManifest(task.taskRoot, task.manifest);
  return pending;
}

function validateReview(review) {
  if (!review || typeof review !== "object" || Array.isArray(review) || review.schemaVersion !== 1) throw taskError("invalid-review", "Commit review must be an OfficeKit review report.");
  if (!new Set(["passed", "passed-with-limitations", "failed"]).has(review.verdict)) throw taskError("invalid-review", "Review verdict is invalid.");
  if (!review.delivery || !isSha(review.delivery.sha256)) throw taskError("invalid-review", "Review delivery SHA-256 is missing.");
  if (!ARTIFACT_KINDS.has(review.artifactKind)) throw taskError("invalid-review", "Review artifact kind is invalid.");
  if (!VISUAL_REVIEW_STATUSES.has(review.visualReview)) throw taskError("invalid-review", "Review visual status is invalid.");
  let encoded;
  try { encoded = JSON.stringify(review); }
  catch (error) { throw taskError("invalid-review", `Review report is not serializable: ${boundedError(error)}`); }
  if (Buffer.byteLength(encoded) > DEFAULT_MAX_REVIEW_REPORT_BYTES) throw taskError("review-too-large", `Review report exceeds ${DEFAULT_MAX_REVIEW_REPORT_BYTES} bytes.`);
  return review;
}

async function writeReviewEvidence(task, name, review) {
  const relative = toPosix(path.join("evidence", "reviews", `${safeName(name)}.json`));
  const target = path.join(task.taskRoot, relative);
  const bytes = Buffer.from(`${JSON.stringify(review, null, 2)}\n`);
  await ensurePrivateDirectory(path.dirname(target), task.taskRoot);
  await atomicWrite(target, bytes, task.taskRoot, { replace: false });
  return { path: relative, bytes: bytes.byteLength, sha256: sha256(bytes) };
}

function reviewSummary(review, evidence) {
  const limitations = [];
  if (review.verdict === "passed-with-limitations") {
    if (review.visualReview !== "complete") limitations.push(`visualReview:${review.visualReview}`);
    if (review.contentView?.requested && review.contentView.status !== "ready") limitations.push(`contentView:${review.contentView.status}`);
    for (const section of ["semantic", "structural", "layout", "design", "delivery"]) {
      const status = review[section]?.status;
      if (status && !new Set(["passed", "ready"]).has(status)) limitations.push(`${section}:${status}`);
    }
  }
  return {
    verdict: review.verdict,
    artifactKind: review.artifactKind,
    format: review.format,
    visualReview: review.visualReview,
    contentView: review.contentView?.requested ? review.contentView.status : "not-requested",
    planSha256: review.design?.planSha256 ?? null,
    deliverySha256: review.delivery.sha256,
    limitations: [...new Set(limitations)],
    evidence,
  };
}

function artifactDescriptor(artifact, taskRoot) {
  return Object.freeze({
    artifactId: artifact.id,
    name: artifact.name,
    kind: artifact.kind,
    mime: artifact.mime,
    path: resolveManagedFile(taskRoot, artifact.source.storedPath, "input"),
    sourcePath: artifact.source.path,
    bytes: artifact.source.bytes,
    sha256: artifact.source.sha256,
  });
}

function deriveTaskState(manifest) {
  if (manifest.pending.length > 0) return "attention";
  if (manifest.plan && planDescriptorForManifest(manifest).state === "working") return "working";
  const ppjHead = manifest.ppj?.revisions.find((revision) => revision.sha256 === manifest.ppj.head);
  if (ppjHead?.status === "review-failed") return "attention";
  if (ppjHead && ppjHead.status !== "reviewed") return "working";
  if (ppjHead?.status === "reviewed") return "stable";
  if (manifest.publications.length > 0) return "published";
  if (manifest.head) return "stable";
  return "new";
}

function taskStorageBytes(manifest) {
  let total = 0;
  for (const artifact of manifest.artifacts) {
    total += artifact.source?.bytes ?? 0;
    total += artifact.headRevision?.bytes ?? 0;
  }
  for (const pending of manifest.pending) total += pending.bytes ?? 0;
  for (const commit of manifest.commits) total += commit.operation?.bytes ?? 0;
  total += manifest.plan?.bytes ?? 0;
  for (const revision of manifest.ppj?.revisions ?? []) {
    total += revision.bytes ?? 0;
    total += revision.resources?.reduce((sum, resource) => sum + (resource.bytes ?? 0), 0) ?? 0;
    total += revision.nodeMap?.bytes ?? 0;
    total += revision.receipts?.reduce((sum, receipt) => sum + (receipt.bytes ?? 0), 0) ?? 0;
    total += revision.candidate?.bytes ?? 0;
    total += revision.review?.evidence?.bytes ?? 0;
  }
  return total;
}

async function storePpjResources(task, workspace, program, revisionRootRelative) {
  const output = [];
  const occupied = new Set(["program.ppj", "node-map.json"]);
  const assets = new Map((workspace.assets ?? []).map((asset) => [asset.id, asset]));
  for (const declaration of program.assets ?? []) {
    const asset = assets.get(declaration.id);
    if (!asset || !(asset.data instanceof Uint8Array) ||
        sha256(asset.data) !== declaration.sha256 || asset.mimeType !== declaration.mimeType) {
      throw taskError("invalid-ppj-resource", `PPJ asset ${declaration.id} is missing or stale while recording the task revision.`);
    }
    output.push(await storePpjResource(task, revisionRootRelative, occupied, {
      kind: "asset",
      id: declaration.id,
      uri: declaration.uri,
      mimeType: declaration.mimeType,
      sha256: declaration.sha256,
      data: asset.data,
    }));
  }
  if (program.source != null) {
    if (!(workspace.source instanceof Uint8Array) || workspace.source.byteLength === 0 ||
        sha256(workspace.source) !== program.source.sha256) {
      throw taskError("invalid-ppj-resource", "PPJ source package is missing or stale while recording the task revision.");
    }
    output.push(await storePpjResource(task, revisionRootRelative, occupied, {
      kind: "source",
      id: null,
      uri: program.source.uri,
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      sha256: program.source.sha256,
      data: workspace.source,
    }));
  }
  return output.sort((left, right) => left.path.localeCompare(right.path));
}

async function storePpjResource(task, revisionRootRelative, occupied, resource) {
  const uri = safePpjTaskUri(resource.uri);
  if (!occupied.add(uri)) throw taskError("invalid-ppj-resource", `PPJ resource URI collides inside the task revision: ${uri}`);
  const relative = `${revisionRootRelative}/${uri}`;
  const target = path.join(task.taskRoot, ...relative.split("/"));
  await ensurePrivateSubdirectory(task.taskRoot, path.dirname(relative));
  const bytes = Buffer.from(resource.data);
  if (bytes.byteLength === 0 || bytes.byteLength > DEFAULT_MAX_TASK_ARTIFACT_BYTES || sha256(bytes) !== resource.sha256) {
    throw taskError("invalid-ppj-resource", `PPJ resource ${uri} exceeds its task budget or has a stale hash.`);
  }
  await writeImmutable(target, bytes, 0o400, { allowIdentical: true });
  return {
    kind: resource.kind,
    id: resource.id,
    uri,
    mimeType: resource.mimeType,
    path: toPosix(relative),
    bytes: bytes.byteLength,
    sha256: resource.sha256,
  };
}

async function storePpjCandidate(task, receipt, candidate) {
  if (!(candidate.bytes instanceof Uint8Array)) throw taskError("invalid-ppj-candidate", "PPJ task candidate bytes are missing.");
  const bytes = Buffer.from(candidate.bytes);
  const digest = sha256(bytes);
  if (bytes.byteLength === 0 || bytes.byteLength > DEFAULT_MAX_TASK_ARTIFACT_BYTES ||
      !isSha(receipt.outputSha256) || digest !== receipt.outputSha256) {
    throw taskError("invalid-ppj-candidate", "PPJ task candidate does not match the native build receipt.");
  }
  const relative = toPosix(path.join("candidates", "ppj", receipt.programSha256, `${digest}.pptx`));
  await ensurePrivateSubdirectory(task.taskRoot, path.dirname(relative));
  await writeImmutable(path.join(task.taskRoot, relative), bytes, 0o400, { allowIdentical: true });
  return {
    path: relative,
    bytes: bytes.byteLength,
    sha256: digest,
    outputPath: typeof candidate.outputPath === "string" ? path.resolve(candidate.outputPath) : null,
  };
}

async function storePpjReview(task, receipt, candidate, review) {
  if (!review || typeof review !== "object" || Array.isArray(review) ||
      !new Set(["passed", "passed-with-limitations", "failed"]).has(review.verdict) ||
      review.delivery?.sha256 !== candidate?.sha256) {
    throw taskError("invalid-ppj-review", "PPJ review must bind the exact task candidate and contain a valid verdict.");
  }
  let encoded;
  try { encoded = Buffer.from(`${JSON.stringify(review, null, 2)}\n`); }
  catch (error) { throw taskError("invalid-ppj-review", `PPJ review is not serializable: ${boundedError(error)}`); }
  if (encoded.byteLength > DEFAULT_MAX_REVIEW_REPORT_BYTES) throw taskError("invalid-ppj-review", "PPJ review exceeds its task budget.");
  const digest = sha256(encoded);
  const relative = toPosix(path.join("evidence", "ppj", receipt.programSha256, `review-${digest}.json`));
  await ensurePrivateSubdirectory(task.taskRoot, path.dirname(relative));
  await writeImmutable(path.join(task.taskRoot, relative), encoded, 0o400, { allowIdentical: true });
  return {
    verdict: review.verdict,
    visualReview: review.visualReview ?? "unavailable",
    playbackEvidence: review.playbackEvidence ?? "structural",
    candidateSha256: candidate.sha256,
    evidence: { path: relative, bytes: encoded.byteLength, sha256: digest },
  };
}

async function verifyPpjTaskRevision(taskRoot, revision) {
  const programPath = resolveManagedFile(taskRoot, revision.path, "PPJ revision");
  const program = await readRegularBounded(programPath, DEFAULT_MAX_TASK_PPJ_BYTES, "PPJ revision");
  if (program.byteLength !== revision.bytes || sha256(program) !== revision.sha256) {
    throw taskError("ppj-revision-corrupt", `PPJ revision ${revision.sha256} failed its program hash.`);
  }
  for (const resource of revision.resources) {
    const target = resolveManagedFile(taskRoot, resource.path, "PPJ resource");
    const bytes = await readRegularBounded(target, DEFAULT_MAX_TASK_ARTIFACT_BYTES, "PPJ resource");
    if (bytes.byteLength !== resource.bytes || sha256(bytes) !== resource.sha256) {
      throw taskError("ppj-revision-corrupt", `PPJ resource ${resource.uri} failed its content hash.`);
    }
  }
  if (revision.nodeMap) await verifyManagedRecord(taskRoot, revision.nodeMap, DEFAULT_MAX_TASK_PPJ_BYTES, "PPJ node map");
  for (const receipt of revision.receipts) await verifyManagedRecord(taskRoot, receipt, DEFAULT_MAX_TASK_MANIFEST_BYTES, "PPJ receipt");
  if (revision.candidate) await verifyManagedRecord(taskRoot, revision.candidate, DEFAULT_MAX_TASK_ARTIFACT_BYTES, "PPJ candidate");
  if (revision.review) await verifyManagedRecord(taskRoot, revision.review.evidence, DEFAULT_MAX_REVIEW_REPORT_BYTES, "PPJ review");
}

async function verifyManagedRecord(taskRoot, record, maximum, label) {
  const target = resolveManagedFile(taskRoot, record.path, label);
  const bytes = await readRegularBounded(target, maximum, label);
  if (bytes.byteLength !== record.bytes || sha256(bytes) !== record.sha256) throw taskError("ppj-revision-corrupt", `${label} failed its content hash.`);
}

function ppjRevisionDescriptor(manifest, taskRoot, { detailed = false, revisionSha256 = manifest.ppj?.head } = {}) {
  const ppj = manifest.ppj;
  if (ppj == null || revisionSha256 == null) return null;
  const revision = ppj.revisions.find((item) => item.sha256 === revisionSha256);
  if (!revision) return null;
  return {
    schema: ppj.schema,
    status: revision.status,
    sha256: revision.sha256,
    mode: revision.identity.mode,
    sourceSha256: revision.identity.sourceSha256,
    bytes: revision.bytes,
    path: detailed && taskRoot ? resolveManagedFile(taskRoot, revision.path, "PPJ revision") : revision.path,
    reviewedSha256: ppj.reviewed,
    candidate: revision.candidate ? { sha256: revision.candidate.sha256, bytes: revision.candidate.bytes } : null,
    review: revision.review ? {
      verdict: revision.review.verdict,
      visualReview: revision.review.visualReview,
      playbackEvidence: revision.review.playbackEvidence,
    } : null,
    updatedAt: revision.updatedAt,
  };
}

function safePpjTaskUri(value) {
  if (typeof value !== "string" || value === "" || value.includes("\\") || value.includes("\0") || value.startsWith("/") ||
      /^[A-Za-z][A-Za-z0-9+.-]*:/u.test(value)) throw taskError("invalid-ppj-resource", "PPJ resource URI must be relative.");
  const segments = value.split("/");
  if (segments.some((segment) => segment === "" || segment === "." || segment === "..")) {
    throw taskError("invalid-ppj-resource", "PPJ resource URI contains an unsafe path segment.");
  }
  return segments.join("/");
}

function validateTaskEditPlan(value, outputSha256) {
  if (value == null) return null;
  if (!value || typeof value !== "object" || Array.isArray(value) || value.schema !== "office-kit/pptx-edit-plan/v1") {
    throw taskError("invalid-edit-plan", "Artifact Edit Plan metadata is invalid.");
  }
  if (!Array.isArray(value.operations) || value.operations.length === 0) return null;
  if (!isSha(value.sourceRevisionSha256) || !isSha(value.outputSha256) || value.outputSha256 !== outputSha256) {
    throw taskError("invalid-edit-plan", "Artifact Edit Plan revisions do not match the committed candidate.");
  }
  if (!Array.isArray(value.changedParts) || value.changedParts.length === 0 || value.changedParts.some((partPath) => !safeChangedOperationPartPath(partPath))) {
    throw taskError("invalid-edit-plan", "Artifact Edit Plan changed parts are invalid.");
  }
  if (new Set(value.changedParts).size !== value.changedParts.length) {
    throw taskError("invalid-edit-plan", "Artifact Edit Plan changed parts must be unique.");
  }
  const changedParts = new Set(value.changedParts);
  const requiredChangedParts = new Set();
  const operationIds = new Set();
  const nativeLeafKinds = new Set([
    "text", "tableCellText", "fillRgb", "lineRgb", "leftEmu", "topEmu", "widthEmu", "heightEmu",
    "imageAsset", "chartTitleText", "chartDataValue", "diagramText", "deleteElement",
  ]);
  for (const operation of value.operations) {
    if (!operation || typeof operation !== "object" || Array.isArray(operation)) {
      throw taskError("invalid-edit-plan", "Artifact Edit Plan operation is invalid.");
    }
    const leafKind = operation.leafKind == null || operation.leafKind === "" ? "text" : operation.leafKind;
    if (!taskOperationKeys(operation, leafKind)) {
      throw taskError("invalid-edit-plan", "Artifact Edit Plan operation contains an undeclared field.");
    }
    if (
        typeof operation.operationId !== "string" || !operation.operationId || operation.operationId.length > 160 || operationIds.has(operation.operationId) ||
        typeof operation.slideId !== "string" || !operation.slideId || typeof operation.targetId !== "string" || !operation.targetId ||
        !safeSlideOperationPartPath(operation.slidePartPath) || !Number.isSafeInteger(operation.shapeTreeIndex) || operation.shapeTreeIndex < 0 ||
        !safeOperationShapeTreePath(operation.shapeTreePath, operation.shapeTreeIndex) ||
        !Number.isSafeInteger(operation.textLeafIndex) || operation.textLeafIndex < 0 ||
        !isSha(operation.expectedSlideSha256) || !isSha(operation.expectedElementSha256) ||
        !isSha(operation.expectedSemanticSha256) || !isSha(operation.expectedTextSha256) ||
        !nativeLeafKinds.has(leafKind) ||
        typeof operation.expectedValue !== "string" || typeof operation.value !== "string" || operation.expectedValue === operation.value) {
      throw taskError("invalid-edit-plan", "Artifact Edit Plan operation is invalid.");
    }
    const footprint = operation.footprint;
    if (!footprint || !exactObjectKeys(footprint, TASK_FOOTPRINT_KEYS) ||
        !isSha(footprint.sourceElementSha256) ||
        !(isSha(footprint.outputElementSha256) || (leafKind === "deleteElement" && footprint.outputElementSha256 === "")) ||
        !isSha(footprint.oldValueSha256) || !isSha(footprint.newValueSha256) ||
        !safeOperationShapeTreePath(footprint.shapeTreePath, operation.shapeTreeIndex) ||
        ((footprint.leafKind == null || footprint.leafKind === "" ? "text" : footprint.leafKind) !== leafKind) ||
        !safeOperationPartPath(footprint.mutationPartPath) || !changedParts.has(footprint.mutationPartPath) ||
        !decimalOffset(footprint.sourceStartOffset) || !decimalOffset(footprint.sourceEndOffset) || !decimalOffset(footprint.outputEndOffset) ||
        !Array.isArray(footprint.nestedFootprints)) {
      throw taskError("invalid-edit-plan", "Artifact Edit Plan mutation footprint is invalid.");
    }
    requiredChangedParts.add(footprint.mutationPartPath);
    validateDependentTaskOperation(operation, footprint, leafKind, changedParts, requiredChangedParts);
    operationIds.add(operation.operationId);
  }
  if (requiredChangedParts.size !== changedParts.size || [...changedParts].some((partPath) => !requiredChangedParts.has(partPath))) {
    throw taskError("invalid-edit-plan", "Artifact Edit Plan changed parts do not match its mutation footprints.");
  }
  let encoded;
  try { encoded = JSON.stringify(value); }
  catch (error) { throw taskError("invalid-edit-plan", `Artifact Edit Plan is not serializable: ${boundedError(error)}`); }
  if (Buffer.byteLength(encoded) > DEFAULT_MAX_TASK_OPERATION_BYTES) {
    throw taskError("edit-plan-too-large", `Artifact Edit Plan exceeds ${DEFAULT_MAX_TASK_OPERATION_BYTES} bytes.`);
  }
  return structuredClone(value);
}

function safeOperationShapeTreePath(value, rootIndex) {
  if (value == null) return true;
  return Array.isArray(value) && value.length > 0 && value.length <= 32 &&
    value[0] === rootIndex && value.every((item) => Number.isSafeInteger(item) && item >= 0 && item <= 4_294_967_295);
}

function safeOperationPartPath(value) {
  return typeof value === "string" &&
    /^ppt\/(?:[A-Za-z0-9_.-]+\/)*[A-Za-z0-9_.-]+[.](?:xml|xlsx)$/iu.test(value) &&
    !value.includes("..");
}

function safeChangedOperationPartPath(value) {
  return safeOperationPartPath(value) || safeSlideRelationshipPartPath(value) || safeOfficeKitMediaPartPath(value);
}

function safeSlideRelationshipPartPath(value) {
  return typeof value === "string" && /^ppt\/slides\/_rels\/slide[1-9][0-9]*[.]xml[.]rels$/iu.test(value) && !value.includes("..");
}

function safeOfficeKitMediaPartPath(value) {
  return typeof value === "string" && /^ppt\/media\/office-kit-[0-9a-f]{24}[.](?:png|jpe?g|gif|svg)$/iu.test(value) && !value.includes("..");
}

function safeSlideOperationPartPath(value) {
  return typeof value === "string" && /^ppt\/slides\/slide[1-9][0-9]*[.]xml$/iu.test(value) && !value.includes("..");
}

function safeEmbeddedWorksheetPartPath(value) {
  return typeof value === "string" && /^xl\/worksheets\/[A-Za-z0-9_.-]+[.]xml$/iu.test(value) && !value.includes("..");
}

const TASK_OPERATION_COMMON_KEYS = new Set([
  "operationId", "slideId", "slidePartPath", "expectedSlideSha256", "targetId", "shapeTreeIndex", "shapeTreePath",
  "leafKind", "expectedElementSha256", "expectedSemanticSha256", "textLeafIndex", "expectedTextSha256",
  "expectedValue", "value", "footprint",
]);
const TASK_OPERATION_DEPENDENT_KEYS = ["targetPartPath", "expectedTargetPartSha256", "relationshipId"];
const TASK_OPERATION_CHART_DATA_KEYS = [
  "embeddedPackagePartPath", "expectedEmbeddedPackageSha256", "embeddedPackageRelationshipId",
  "embeddedWorksheetPartPath", "expectedEmbeddedWorksheetSha256", "embeddedCellReference",
  "chartSeriesIndex", "chartPointIndex", "chartFormula",
];
const TASK_OPERATION_DIAGRAM_KEYS = ["diagramModelId", "diagramRunIndex"];
const TASK_OPERATION_IMAGE_KEYS = ["imageReplacement"];
const TASK_OPERATION_DELETION_KEYS = ["elementDeletion"];
const TASK_IMAGE_REPLACEMENT_KEYS = new Set([
  "assetId", "sha256", "contentType", "byteLength", "relationshipPartPath", "mediaPartPath", "crop",
]);
const TASK_IMAGE_CROP_KEYS = new Set([
  "leftThousandthPercent", "topThousandthPercent", "rightThousandthPercent", "bottomThousandthPercent",
]);
const TASK_ELEMENT_DELETION_KEYS = new Set(["expectedNativeId"]);
const TASK_IMAGE_EXTENSIONS_BY_CONTENT_TYPE = new Map([
  ["image/png", new Set(["png"])],
  ["image/jpeg", new Set(["jpg", "jpeg"])],
  ["image/gif", new Set(["gif"])],
  ["image/svg+xml", new Set(["svg"])],
]);
const TASK_FOOTPRINT_KEYS = new Set([
  "mutationPartPath", "sourceElementSha256", "outputElementSha256", "oldValueSha256", "newValueSha256",
  "sourceStartOffset", "sourceEndOffset", "outputEndOffset", "shapeTreePath", "leafKind", "nestedFootprints",
]);
const TASK_NESTED_FOOTPRINT_KEYS = new Set([
  "containerPartPath", "partPath", "oldValueSha256", "newValueSha256", "sourceStartOffset", "sourceEndOffset", "outputEndOffset",
]);

function taskOperationKeys(operation, leafKind) {
  const allowed = new Set(TASK_OPERATION_COMMON_KEYS);
  if (leafKind === "chartTitleText" || leafKind === "chartDataValue" || leafKind === "diagramText") {
    for (const key of TASK_OPERATION_DEPENDENT_KEYS) allowed.add(key);
  }
  if (leafKind === "chartDataValue") for (const key of TASK_OPERATION_CHART_DATA_KEYS) allowed.add(key);
  if (leafKind === "diagramText") for (const key of TASK_OPERATION_DIAGRAM_KEYS) allowed.add(key);
  if (leafKind === "imageAsset") for (const key of TASK_OPERATION_IMAGE_KEYS) allowed.add(key);
  if (leafKind === "deleteElement") for (const key of TASK_OPERATION_DELETION_KEYS) allowed.add(key);
  return exactObjectKeys(operation, allowed);
}

function exactObjectKeys(value, allowed) {
  return value && typeof value === "object" && !Array.isArray(value) && Object.keys(value).every((key) => allowed.has(key));
}

function validateDependentTaskOperation(operation, footprint, leafKind, changedParts, requiredChangedParts) {
  if (leafKind === "imageAsset") {
    validateTaskImageOperation(operation, footprint, changedParts, requiredChangedParts);
    return;
  }
  if (leafKind === "deleteElement") {
    const deletion = operation.elementDeletion;
    if (footprint.mutationPartPath !== operation.slidePartPath || footprint.nestedFootprints.length !== 0 ||
        !deletion || !exactObjectKeys(deletion, TASK_ELEMENT_DELETION_KEYS) ||
        !Number.isSafeInteger(deletion.expectedNativeId) || deletion.expectedNativeId <= 0 || deletion.expectedNativeId > 4_294_967_295 ||
        operation.shapeTreePath?.length !== 1 || operation.expectedValue !== operation.targetId || operation.value !== "") {
      throw taskError("invalid-edit-plan", "Element-deletion Edit Plan binding is invalid.");
    }
    return;
  }
  const dependent = leafKind === "chartTitleText" || leafKind === "chartDataValue" || leafKind === "diagramText";
  if (!dependent) {
    if (footprint.mutationPartPath !== operation.slidePartPath || footprint.nestedFootprints.length !== 0 ||
        operation.targetPartPath != null || operation.embeddedPackagePartPath != null || operation.diagramModelId != null) {
      throw taskError("invalid-edit-plan", "Slide Edit Plan operation contains a foreign dependent-part binding.");
    }
    return;
  }
  if (!safeXmlDependentPartPath(operation.targetPartPath) || !isSha(operation.expectedTargetPartSha256) ||
      !boundedIdentifier(operation.relationshipId, 255) || footprint.mutationPartPath !== operation.targetPartPath) {
    throw taskError("invalid-edit-plan", "Dependent Edit Plan operation binding is invalid.");
  }
  if (leafKind === "chartDataValue") {
    if (!safeChartPartPath(operation.targetPartPath) || !finiteNumericToken(operation.expectedValue) || !finiteNumericToken(operation.value) ||
        !safeEmbeddedPackagePartPath(operation.embeddedPackagePartPath) || !isSha(operation.expectedEmbeddedPackageSha256) ||
        !boundedIdentifier(operation.embeddedPackageRelationshipId, 255) ||
        !safeEmbeddedWorksheetPartPath(operation.embeddedWorksheetPartPath) || !isSha(operation.expectedEmbeddedWorksheetSha256) ||
        typeof operation.embeddedCellReference !== "string" || !/^[A-Z]{1,3}[1-9][0-9]*$/u.test(operation.embeddedCellReference) ||
        !Number.isSafeInteger(operation.chartSeriesIndex) || operation.chartSeriesIndex < 0 ||
        !Number.isSafeInteger(operation.chartPointIndex) || operation.chartPointIndex < 0 ||
        typeof operation.chartFormula !== "string" || operation.chartFormula.length === 0 || operation.chartFormula.length > 32_768 ||
        footprint.nestedFootprints.length !== 1) {
      throw taskError("invalid-edit-plan", "Chart-data Edit Plan binding is invalid.");
    }
    const [nested] = footprint.nestedFootprints;
    if (!nested || !exactObjectKeys(nested, TASK_NESTED_FOOTPRINT_KEYS) || nested.containerPartPath !== operation.embeddedPackagePartPath ||
        nested.partPath !== operation.embeddedWorksheetPartPath || !isSha(nested.oldValueSha256) || !isSha(nested.newValueSha256) ||
        !decimalOffset(nested.sourceStartOffset) || !decimalOffset(nested.sourceEndOffset) || !decimalOffset(nested.outputEndOffset) ||
        !changedParts.has(nested.containerPartPath)) {
      throw taskError("invalid-edit-plan", "Chart-data Edit Plan nested footprint is invalid.");
    }
    requiredChangedParts.add(nested.containerPartPath);
    if (operation.diagramModelId != null || operation.diagramRunIndex != null) {
      throw taskError("invalid-edit-plan", "Chart-data Edit Plan operation contains a foreign SmartArt binding.");
    }
    return;
  }
  if (footprint.nestedFootprints.length !== 0 || operation.embeddedPackagePartPath != null ||
      operation.embeddedWorksheetPartPath != null || operation.chartFormula != null) {
    throw taskError("invalid-edit-plan", "Dependent XML Edit Plan operation contains a foreign embedded-package binding.");
  }
  if (leafKind === "diagramText") {
    if (!safeDiagramPartPath(operation.targetPartPath) || !boundedIdentifier(operation.diagramModelId, 1_024) ||
        !Number.isSafeInteger(operation.diagramRunIndex) || operation.diagramRunIndex < 0) {
      throw taskError("invalid-edit-plan", "SmartArt Edit Plan binding is invalid.");
    }
  } else if (!safeChartPartPath(operation.targetPartPath) || operation.diagramModelId != null || operation.diagramRunIndex != null) {
    throw taskError("invalid-edit-plan", "Chart-title Edit Plan operation contains a foreign SmartArt binding.");
  }
}

function validateTaskImageOperation(operation, footprint, changedParts, requiredChangedParts) {
  const replacement = operation.imageReplacement;
  if (footprint.mutationPartPath !== operation.slidePartPath || footprint.nestedFootprints.length !== 0 ||
      !replacement || !exactObjectKeys(replacement, TASK_IMAGE_REPLACEMENT_KEYS) ||
      replacement.assetId !== operation.value || !isSha(replacement.sha256) ||
      replacement.assetId !== `asset/presentation/picture-bullet/${replacement.sha256}` ||
      !Number.isSafeInteger(replacement.byteLength) || replacement.byteLength <= 0 || replacement.byteLength > 16 * 1024 * 1024) {
    throw taskError("invalid-edit-plan", "Image Edit Plan binding is invalid.");
  }
  const relationshipPartPath = slideRelationshipPartPath(operation.slidePartPath);
  if (!relationshipPartPath || replacement.relationshipPartPath !== relationshipPartPath || !changedParts.has(relationshipPartPath)) {
    throw taskError("invalid-edit-plan", "Image Edit Plan relationship footprint is invalid.");
  }
  const extensions = TASK_IMAGE_EXTENSIONS_BY_CONTENT_TYPE.get(replacement.contentType);
  if (!extensions) throw taskError("invalid-edit-plan", "Image Edit Plan content type is invalid.");
  const mediaPrefix = `ppt/media/office-kit-${replacement.sha256.slice(0, 24)}.`;
  if (replacement.mediaPartPath !== null) {
    const extension = typeof replacement.mediaPartPath === "string" && replacement.mediaPartPath.startsWith(mediaPrefix)
      ? replacement.mediaPartPath.slice(mediaPrefix.length)
      : "";
    if (!extensions.has(extension) || !safeOfficeKitMediaPartPath(replacement.mediaPartPath) || !changedParts.has(replacement.mediaPartPath)) {
      throw taskError("invalid-edit-plan", "Image Edit Plan media footprint is invalid.");
    }
    requiredChangedParts.add(replacement.mediaPartPath);
  }
  if (replacement.crop !== null && !validTaskImageCrop(replacement.crop)) {
    throw taskError("invalid-edit-plan", "Image Edit Plan crop is invalid.");
  }
  requiredChangedParts.add(relationshipPartPath);
}

function slideRelationshipPartPath(slidePartPath) {
  const match = String(slidePartPath).match(/^ppt\/slides\/(slide[1-9][0-9]*[.]xml)$/u);
  return match ? `ppt/slides/_rels/${match[1]}.rels` : undefined;
}

function validTaskImageCrop(value) {
  if (!exactObjectKeys(value, TASK_IMAGE_CROP_KEYS)) return false;
  const edges = [value.leftThousandthPercent, value.topThousandthPercent, value.rightThousandthPercent, value.bottomThousandthPercent];
  if (edges.some((edge) => !Number.isSafeInteger(edge) || edge < -100_000 || edge > 100_000)) return false;
  return value.leftThousandthPercent + value.rightThousandthPercent < 100_000 &&
    value.topThousandthPercent + value.bottomThousandthPercent < 100_000;
}

function safeXmlDependentPartPath(value) {
  return typeof value === "string" && /^ppt\/(?:[A-Za-z0-9_.-]+\/)*[A-Za-z0-9_.-]+[.]xml$/iu.test(value) && !value.includes("..");
}

function safeChartPartPath(value) {
  return typeof value === "string" && /^ppt\/charts\/[A-Za-z0-9_.-]+[.]xml$/iu.test(value) && !value.includes("..");
}

function safeDiagramPartPath(value) {
  return typeof value === "string" && /^ppt\/diagrams\/[A-Za-z0-9_.-]+[.]xml$/iu.test(value) && !value.includes("..");
}

function safeEmbeddedPackagePartPath(value) {
  return typeof value === "string" && /^ppt\/embeddings\/[A-Za-z0-9_.-]+[.]xlsx$/iu.test(value) && !value.includes("..");
}

function boundedIdentifier(value, maximum) {
  return typeof value === "string" && value.length > 0 && value.length <= maximum && !/[\u0000-\u001F\u007F]/u.test(value);
}

function finiteNumericToken(value) {
  return typeof value === "string" && value.length <= 128 &&
    /^-?(?:0|[1-9][0-9]*)(?:[.][0-9]+)?(?:[Ee][+-]?[0-9]+)?$/u.test(value) && Number.isFinite(Number(value));
}

function decimalOffset(value) {
  return typeof value === "string" && /^(0|[1-9][0-9]*)$/u.test(value);
}

async function writeTaskOperationRecord(task, commitId, artifactId, plan) {
  const record = {
    schema: "office-kit/task-edit-plan/v1",
    taskId: task.manifest.id,
    commitId,
    artifactId,
    plan,
  };
  const bytes = Buffer.from(`${JSON.stringify(record, null, 2)}\n`);
  if (bytes.byteLength > DEFAULT_MAX_TASK_OPERATION_BYTES) {
    throw taskError("edit-plan-too-large", `Task Edit Plan record exceeds ${DEFAULT_MAX_TASK_OPERATION_BYTES} bytes.`);
  }
  const digest = sha256(bytes);
  const relative = toPosix(path.join("operations", `${commitId}-${artifactId}-${digest}.json`));
  const target = path.join(task.taskRoot, relative);
  await writeImmutable(target, bytes, 0o400, { allowIdentical: true });
  return {
    schema: record.schema,
    path: relative,
    bytes: bytes.byteLength,
    sha256: digest,
    sourceRevisionSha256: plan.sourceRevisionSha256,
    outputRevisionSha256: plan.outputSha256,
    operationCount: plan.operations.length,
    changedParts: [...plan.changedParts],
  };
}

function validateTaskOperationSummary(summary) {
  if (!summary || summary.schema !== "office-kit/task-edit-plan/v1" || typeof summary.path !== "string" ||
      !Number.isSafeInteger(summary.bytes) || summary.bytes <= 0 || summary.bytes > DEFAULT_MAX_TASK_OPERATION_BYTES ||
      !isSha(summary.sha256) || !isSha(summary.sourceRevisionSha256) || !isSha(summary.outputRevisionSha256) ||
      !Number.isSafeInteger(summary.operationCount) || summary.operationCount <= 0 ||
      !Array.isArray(summary.changedParts) || summary.changedParts.length === 0 || summary.changedParts.some((partPath) => !safeChangedOperationPartPath(partPath))) {
    throw taskError("invalid-task", "Task operation record metadata is invalid.");
  }
}

async function readTaskOperationRecord(taskRoot, summary) {
  validateTaskOperationSummary(summary);
  const target = resolveManagedFile(taskRoot, summary.path, "operation record");
  const bytes = await readRegularBounded(target, DEFAULT_MAX_TASK_OPERATION_BYTES, "Task operation record");
  if (bytes.byteLength !== summary.bytes || sha256(bytes) !== summary.sha256) {
    throw taskError("operation-corrupt", "Task operation record hash verification failed.");
  }
  let record;
  try { record = JSON.parse(bytes.toString("utf8")); }
  catch (error) { throw taskError("operation-corrupt", `Task operation record is not valid JSON: ${boundedError(error)}`); }
  if (record?.schema !== summary.schema || record?.plan?.sourceRevisionSha256 !== summary.sourceRevisionSha256 ||
      record?.plan?.outputSha256 !== summary.outputRevisionSha256 || record?.plan?.operations?.length !== summary.operationCount) {
    throw taskError("operation-corrupt", "Task operation record does not match its manifest binding.");
  }
  validateTaskEditPlan(record.plan, summary.outputRevisionSha256);
  return record;
}

async function validateTaskOperationFiles(manifest, taskRoot) {
  for (const commit of manifest.commits) {
    if (commit.operation) await readTaskOperationRecord(taskRoot, commit.operation);
  }
}

function taskOperationDescriptor(taskRoot, summary, record) {
  return Object.freeze({
    ...structuredClone(summary),
    path: resolveManagedFile(taskRoot, summary.path, "operation record"),
    operationIds: record.plan.operations.map((operation) => operation.operationId),
  });
}

async function readArtifactBytes(value, maximum = DEFAULT_MAX_TASK_ARTIFACT_BYTES) {
  const limit = positiveInteger(maximum, DEFAULT_MAX_TASK_ARTIFACT_BYTES, "maxBytes");
  let bytes;
  if (typeof value === "string") bytes = await readRegularBounded(value, limit, "Artifact candidate");
  else if (value instanceof Uint8Array) bytes = Buffer.from(value);
  else if (value instanceof ArrayBuffer) bytes = Buffer.from(value);
  else if (ArrayBuffer.isView(value)) bytes = Buffer.from(value.buffer, value.byteOffset, value.byteLength);
  else if (typeof value?.arrayBuffer === "function") bytes = Buffer.from(await value.arrayBuffer());
  else throw taskError("invalid-artifact", "Artifact candidate must be a FileBlob, byte array, ArrayBuffer, or regular file path.");
  if (bytes.byteLength > limit) throw taskError("artifact-too-large", `Artifact candidate exceeds ${limit} bytes.`);
  return bytes;
}

async function readRegularBounded(target, maximum, label) {
  const descriptor = await lstatIfExists(target);
  if (!descriptor || descriptor.isSymbolicLink() || !descriptor.isFile()) throw taskError("unsafe-path", `${label} must be a regular non-symlink file.`);
  if (descriptor.size > maximum) throw taskError("file-too-large", `${label} exceeds ${maximum} bytes.`);
  return readFile(target);
}

async function findInterruptedRequest(target) {
  const descriptor = await lstatIfExists(target);
  if (descriptor == null) return null;
  if (descriptor.isSymbolicLink() || !descriptor.isFile() || descriptor.size > 16_777_216) return null;
  const started = new Map();
  for (const line of (await readFile(target, "utf8")).split(/\r?\n/u).filter(Boolean)) {
    let record;
    try { record = JSON.parse(line); } catch { continue; }
    if (record?.type === "request.started" && Number.isInteger(record.sequence)) started.set(record.sequence, record);
    if (record?.type === "request.terminal" && Number.isInteger(record.sequence)) started.delete(record.sequence);
  }
  return [...started.values()].sort((a, b) => a.sequence - b.sequence).at(-1) ?? null;
}

async function ensurePrivateDirectory(target, containmentRoot) {
  const existing = await lstatIfExists(target);
  if (existing?.isSymbolicLink() || (existing && !existing.isDirectory())) throw taskError("unsafe-path", `Managed path must be a regular directory: ${target}`);
  if (!existing) await mkdir(target, { recursive: false, mode: 0o700 });
  const canonical = await realpath(target);
  const root = await realpath(containmentRoot);
  assertContained(canonical, root, "managed directory");
  await privateMode(canonical, 0o700);
  return canonical;
}

async function ensurePrivateSubdirectory(root, relative) {
  if (typeof relative !== "string" || relative === "" || path.isAbsolute(relative) || relative.includes("\\")) {
    throw taskError("unsafe-path", "Managed subdirectory path must be relative.");
  }
  const segments = relative.split("/");
  if (segments.some((segment) => segment === "" || segment === "." || segment === "..")) {
    throw taskError("unsafe-path", "Managed subdirectory path contains an unsafe segment.");
  }
  const canonicalRoot = await realpath(root);
  let current = canonicalRoot;
  for (const segment of segments) {
    const next = path.join(current, segment);
    const existing = await lstatIfExists(next);
    if (existing?.isSymbolicLink() || existing && !existing.isDirectory()) {
      throw taskError("unsafe-path", `Managed path must be a regular directory: ${next}`);
    }
    if (!existing) await mkdir(next, { mode: 0o700 });
    current = await realpath(next);
    assertContained(current, canonicalRoot, "managed subdirectory");
    await privateMode(current, 0o700);
  }
  return current;
}

async function canonicalDirectory(target, label) {
  const requested = path.resolve(target);
  const descriptor = await lstatIfExists(requested);
  if (!descriptor || descriptor.isSymbolicLink() || !descriptor.isDirectory()) throw taskError("invalid-workspace", `${label} must be an existing regular directory: ${requested}`);
  return realpath(requested);
}

async function writeImmutable(target, bytes, mode, { allowIdentical = false } = {}) {
  const existing = await lstatIfExists(target);
  if (existing) {
    if (!allowIdentical || existing.isSymbolicLink() || !existing.isFile()) throw taskError("output-exists", `Managed file already exists: ${target}`);
    const current = await readFile(target);
    if (sha256(current) !== sha256(bytes)) throw taskError("revision-collision", "Immutable revision path contains different bytes.");
    return;
  }
  await atomicWrite(target, bytes, path.dirname(path.dirname(target)), { replace: false, mode });
  await privateMode(target, mode);
}

async function atomicWrite(target, bytes, containmentRoot, { replace, mode = 0o600 }) {
  const parent = await realpath(path.dirname(target));
  const root = await realpath(containmentRoot);
  assertContained(parent, root, "managed file parent");
  const existing = await lstatIfExists(target);
  if (existing?.isSymbolicLink() || (existing && !existing.isFile())) throw taskError("unsafe-path", `Managed file must be a regular file: ${target}`);
  if (existing && !replace) throw taskError("output-exists", `Managed file already exists: ${target}`);
  const temporary = path.join(parent, `.${path.basename(target)}.${randomUUID()}.tmp`);
  await writeFile(temporary, bytes, { mode });
  try { await rename(temporary, target); }
  finally { await rm(temporary, { force: true }); }
  await privateMode(target, mode);
}

function resolveManagedFile(taskRoot, relative, label) {
  if (typeof relative !== "string" || relative === "" || path.isAbsolute(relative)) throw taskError("invalid-task", `Task ${label} path is invalid.`);
  const target = path.resolve(taskRoot, relative);
  assertContained(target, taskRoot, label);
  return target;
}

function assertOutsideManagedTasks(candidate, workspaceRoot) {
  const tasksRoot = path.join(workspaceRoot, TASK_DIRECTORY);
  const relative = path.relative(tasksRoot, candidate);
  if (relative === "" || (!relative.startsWith(`..${path.sep}`) && relative !== ".." && !path.isAbsolute(relative))) {
    throw taskError("unsafe-input", "A task input cannot be staged from managed task state.");
  }
}

function assertContained(candidate, root, label) {
  const relative = path.relative(root, candidate);
  if (relative === ".." || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) throw taskError("unsafe-path", `${label} escapes its managed root.`);
}

function validateTaskId(value) {
  if (typeof value !== "string" || !TASK_ID_PATTERN.test(value)) throw taskError("invalid-task-id", "Task ID is invalid.");
  return value;
}

function validateArtifactId(value) {
  if (typeof value !== "string" || !ARTIFACT_ID_PATTERN.test(value)) throw taskError("invalid-artifact-id", "Artifact ID must be a lowercase safe identifier of at most 64 characters.");
  return value;
}

function newArtifactId() {
  return `a_${randomUUID().replaceAll("-", "").slice(0, 12)}`;
}

function normalizeKind(value, fileName = "", mime = "") {
  if (ARTIFACT_KINDS.has(value)) return value;
  const lowerMime = String(mime).toLowerCase();
  if (lowerMime === "application/pdf") return "pdf";
  const extension = path.extname(String(fileName || "")).toLowerCase();
  if (extension === ".docx") return "document";
  if (extension === ".xlsx") return "workbook";
  if (extension === ".pptx") return "presentation";
  if (extension === ".pdf") return "pdf";
  throw taskError("invalid-artifact-kind", "Artifact kind must be document, workbook, presentation, or pdf.");
}

function extensionFor(fileName, kind) {
  const expected = { document: ".docx", workbook: ".xlsx", presentation: ".pptx", pdf: ".pdf" }[kind];
  return path.extname(String(fileName || "")).toLowerCase() === expected ? expected : expected;
}

function mimeForKind(kind) {
  return {
    document: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    workbook: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    presentation: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    pdf: "application/pdf",
  }[kind];
}

function boundedText(value, label, maximum) {
  if (typeof value !== "string" || value.trim() === "" || value.length > maximum || /[\u0000-\u001f\u007f]/u.test(value)) throw taskError("invalid-text", `${label} must be a non-empty bounded text value.`);
  return value.trim();
}

function normalizeConstraints(value) {
  if (!Array.isArray(value) || value.length > 32) throw taskError("invalid-constraints", "Commit constraints must be an array of at most 32 strings.");
  return [...new Set(value.map((entry) => boundedText(entry, "Commit constraint", 512)))];
}

function normalizeListLimit(value) {
  return positiveInteger(value, DEFAULT_TASK_LIST_LIMIT, "limit");
}

function positiveInteger(value, fallback, label) {
  if (value == null) return fallback;
  const number = Number(value);
  if (!Number.isSafeInteger(number) || number <= 0) throw taskError("invalid-limit", `${label} must be a positive safe integer.`);
  return number;
}

function isSha(value) {
  return typeof value === "string" && /^[a-f0-9]{64}$/u.test(value);
}

function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

function safeName(value) {
  return String(value).replace(/[^a-zA-Z0-9._-]/gu, "-").slice(0, 160);
}

function toPosix(value) {
  return value.split(path.sep).join("/");
}

async function lstatIfExists(target) {
  try { return await lstat(target); }
  catch (error) {
    if (error?.code === "ENOENT") return null;
    throw error;
  }
}

async function readSmallJson(target) {
  const descriptor = await lstatIfExists(target);
  if (!descriptor || descriptor.isSymbolicLink() || !descriptor.isFile() || descriptor.size > 16_384) throw taskError("unsafe-lock", "Task lock is invalid.");
  return JSON.parse(await readFile(target, "utf8"));
}

function processExists(pid) {
  if (!Number.isSafeInteger(pid) || pid <= 0) return false;
  try { process.kill(pid, 0); return true; }
  catch (error) { return error?.code === "EPERM"; }
}

async function privateMode(target, mode) {
  if (process.platform !== "win32") await chmod(target, mode);
}

async function directoryBytes(root) {
  let total = 0;
  for (const entry of await readdir(root, { withFileTypes: true })) {
    const target = path.join(root, entry.name);
    if (entry.isSymbolicLink()) throw taskError("unsafe-task", "Task directories cannot contain symbolic links.");
    if (entry.isDirectory()) total += await directoryBytes(target);
    else if (entry.isFile()) total += (await stat(target)).size;
    else throw taskError("unsafe-task", "Task directories can contain only regular files and directories.");
  }
  return total;
}

function boundedError(error) {
  return String(error?.message || error || "invalid task").replace(/[\u0000-\u001f\u007f]/gu, " ").slice(0, 500);
}

export function taskError(code, message, details = {}) {
  const error = new Error(message);
  error.code = code;
  Object.assign(error, details);
  return error;
}
