import { createHash, randomUUID } from "node:crypto";
import { once } from "node:events";
import { createInterface } from "node:readline";
import {
  appendFile,
  chmod,
  lstat,
  mkdir,
  readFile,
  realpath,
  rename,
  rm,
  writeFile,
} from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { pathToFileURL } from "node:url";
import {
  createOfficeKitResolver,
  installOfficeKitModuleHooks,
  readOfficeKitPackageMetadata,
  resolveWorkspaceSpecifier,
} from "./officekit-resolver.mjs";
import {
  acquireTaskLock,
  beginTaskSession,
  commitTaskArtifact,
  createTask,
  openTask,
  recordTaskPublication,
  readTaskPlan,
  resolveCommittedArtifact,
  stageTaskInput,
  summarizeTask,
  writeTaskPlan,
} from "./task-store.mjs";

export const REPL_PROTOCOL_VERSION = 3;
export const DEFAULT_MAX_REQUEST_BYTES = 1_048_576;
export const DEFAULT_MAX_RESPONSE_BYTES = 8_388_608;
export const DEFAULT_MAX_EVENT_BYTES = 16_384;
export const DEFAULT_MAX_JOURNAL_BYTES = 16_777_216;
export const REPL_USAGE = [
  "Usage:",
  "  officekit repl --new <goal> [--workspace <path>] [--file <cell.mjs>] [limits]",
  "  officekit repl <task-id> [--workspace <path>] [--file <cell.mjs>] [limits]",
  "",
  "Run a local JSONL JavaScript task with a persistent OfficeKit context.",
  "Each input line must contain {id, code}; stdout contains one JSON response per line.",
  "",
  "Options:",
  "  --new <goal>              Create one durable task with a natural-language goal",
  "  --workspace <path>        Workspace root (default: nearest .office-kit, Git root, or cwd)",
  "  --file <cell.mjs>          Execute one JavaScript cell from a regular UTF-8 file, then exit",
  "  --max-request-bytes <n>   Maximum request line size",
  "  --max-response-bytes <n>  Maximum serialized response size",
  "  --max-journal-bytes <n>   Maximum checkpoint journal size",
  "  --help, -h                Show this help",
].join("\n");

export async function runReplCommand(
  args,
  { input = process.stdin, output = process.stdout, errorOutput = process.stderr } = {},
) {
  const options = parseReplArguments(args);
  if (options.help) {
    output.write(`${REPL_USAGE}\n`);
    return;
  }
  const fileRequest = options.codeFile == null
    ? null
    : await readCodeFileRequest(options.codeFile, options.maxRequestBytes);
  const session = await createReplSession(options);
  let reader;
  try {
    await writeJsonLine(output, session.ready, options.maxResponseBytes);
    if (fileRequest != null) {
      const response = await session.handleLine(fileRequest);
      await writeJsonLine(output, response, options.maxResponseBytes);
    } else {
      reader = createInterface({ input, crlfDelay: Infinity });
      for await (const line of reader) {
        if (line.trim() === "") continue;
        const response = await session.handleLine(line);
        await writeJsonLine(output, response, options.maxResponseBytes);
      }
    }
  } finally {
    reader?.close();
    await session.close();
    if (errorOutput && session.diagnostics.length > 0) {
      for (const diagnostic of session.diagnostics) {
        errorOutput.write(`OfficeKit REPL: ${diagnostic}\n`);
      }
    }
  }
}

export async function createReplSession(options = {}) {
  const parsed = await resolveSessionOptions(options);
  const metadata = await readOfficeKitPackageMetadata();
  const resolver = createOfficeKitResolver(metadata);
  const sessionId = randomUUID();
  const opened = parsed.newTaskGoal != null
    ? await createTask({ workspaceRoot: parsed.workspaceRoot, goal: parsed.newTaskGoal })
    : await openTask({ workspaceRoot: parsed.workspaceRoot, taskId: parsed.taskId });
  const lock = await acquireTaskLock(opened.taskRoot, { sessionId });
  let task;
  let roots;
  try {
    task = await beginTaskSession(opened, { sessionId });
    roots = {
      workspaceRoot: task.workspaceRoot,
      taskRoot: task.taskRoot,
      inputRoot: path.join(task.taskRoot, "inputs"),
      assetRoot: path.join(task.workspaceRoot, "assets"),
      outputRoot: path.join(task.workspaceRoot, "outputs"),
      evidenceRoot: path.join(task.taskRoot, "evidence"),
    };
    roots.outputRoot = await ensureDirectory(roots.outputRoot, "output root");
    assertContainedCanonical(roots.outputRoot, roots.workspaceRoot, "output root");
    roots.evidenceRoot = await ensureDirectory(roots.evidenceRoot, "evidence root");
    assertContainedCanonical(roots.evidenceRoot, roots.taskRoot, "evidence root");
  } catch (error) {
    await lock.release();
    throw error;
  }
  const { taskRoot, workspaceRoot } = roots;
  const sessionRootPath = task.sessionRoot;
  const checkpointPath = path.join(sessionRootPath, "checkpoint.json");
  const journalPath = path.join(sessionRootPath, "session.jsonl");
  const previous = null;
  let interruptedRequest = task.manifest.pending.filter((entry) => entry.type === "interrupted-request").at(-1) ?? null;
  const initialState = Object.create(null);
  const sequence = 0;
  const diagnostics = [];
  const imports = [];
  const artifacts = Array.isArray(previous?.artifacts) ? structuredClone(previous.artifacts) : [];
  const evidence = Array.isArray(previous?.evidence) ? structuredClone(previous.evidence) : [];
  const state = initialState;
  const excelFacade = createLazyExcelFacade();
  const powerpointFacade = createLazyPowerPointFacade();
  let closed = false;
  let currentSequence = sequence;
  const deregisterHooks = installOfficeKitModuleHooks(resolver);

  const ctx = {
    protocol: REPL_PROTOCOL_VERSION,
    sessionId,
    ...roots,
    checkpointRoot: sessionRootPath,
    task: Object.freeze(summarizeTask(task.manifest, { detailed: true, taskRoot: task.taskRoot })),
    state,
    import: async (specifier) => {
      const target = resolveWorkspaceSpecifier(resolver, specifier, { workspaceRoot });
      imports.push(specifier);
      return import(target);
    },
    input: async (sourcePath, inputOptions = {}) => {
      const descriptor = await stageTaskInput(task, sourcePath, inputOptions);
      ctx.task = Object.freeze(summarizeTask(task.manifest, { detailed: true, taskRoot: task.taskRoot }));
      return descriptor;
    },
    plan: async (...args) => {
      if (args.length === 0) return readTaskPlan(task);
      try {
        return await writeTaskPlan(task, args[0], args[1] ?? {});
      } finally {
        ctx.task = Object.freeze(summarizeTask(task.manifest, { detailed: true, taskRoot: task.taskRoot }));
      }
    },
    commit: async (value, commitOptions = {}) => {
      try {
        return await commitTaskArtifact(task, value, commitOptions);
      } finally {
        ctx.task = Object.freeze(summarizeTask(task.manifest, { detailed: true, taskRoot: task.taskRoot }));
      }
    },
    publish: async (commitDescriptor, publishOptions = {}) => {
      const committed = await resolveCommittedArtifact(task, commitDescriptor, publishOptions.artifactId);
      const descriptor = await publishArtifact(committed.bytes, {
        ...publishOptions,
        kind: committed.artifact.kind,
        mime: committed.artifact.mime,
        sourcePaths: committed.artifact.source?.path ? [committed.artifact.source.path] : [],
        visualReview: committed.review.visualReview,
        reviewVerdict: committed.review.verdict,
        contentView: committed.review.contentView,
        limitations: committed.review.limitations,
      }, {
        roots,
        sequence: currentSequence,
        artifacts,
      });
      await recordTaskPublication(task, commitDescriptor, descriptor, committed.artifact.id);
      ctx.task = Object.freeze(summarizeTask(task.manifest, { detailed: true, taskRoot: task.taskRoot }));
      return descriptor;
    },
    recordEvidence: async (target, metadata = {}) => {
      const descriptor = await recordEvidenceFile(target, metadata, {
        evidenceRoot: roots.evidenceRoot,
        evidence,
      });
      return descriptor;
    },
  };
  Object.defineProperty(ctx, "excel", {
    enumerable: true,
    configurable: false,
    get: () => excelFacade,
  });
  Object.defineProperty(ctx, "powerpoint", {
    enumerable: true,
    configurable: false,
    get: () => powerpointFacade,
  });

  const session = {
    ctx,
    diagnostics,
    ready: task.ready,
    async handleLine(line) {
      if (closed) return failureResponse(null, replError("session-closed", "REPL session is closed."));
      const request = parseRequestLine(line, parsed.maxRequestBytes);
      if (!request.ok) {
        const response = failureResponse(request.id, request.error, { maybeApplied: false });
        await persistTerminal({ request: request.raw, response, sequence: currentSequence });
        return response;
      }
      currentSequence += 1;
      const started = {
        protocol: REPL_PROTOCOL_VERSION,
        type: "request.started",
        sessionId,
        sequence: currentSequence,
        id: request.id,
        source: request.code,
        sourceSha256: sha256(request.code),
        at: new Date().toISOString(),
      };
      await appendJournal(started);
      const events = [];
      const scopedConsole = createScopedConsole(events);
      const taskBefore = taskResponseMarker(ctx.task);
      let response;
      try {
        const result = await executeCell(request.code, ctx, scopedConsole);
        const maybeApplied = Boolean(interruptedRequest);
        const taskDelta = taskResponseDelta(taskBefore, ctx.task);
        response = successResponse(request.id, {
          result: serializeValue(result),
          ...taskDelta,
          events: boundedEvents(events),
          imports: [...imports.splice(0)],
          artifacts: [...artifacts],
          evidence: [...evidence],
          audit: {
            sequence: currentSequence,
            sourceSha256: sha256(request.code),
            maybeApplied,
            ...(interruptedRequest ? { interruptedRequest } : {}),
          },
        });
      } catch (error) {
        const taskDelta = taskResponseDelta(taskBefore, ctx.task);
        const maybeApplied = error?.maybeApplied === true ||
          taskDelta.commit != null ||
          taskDelta.publication != null ||
          (error?.maybeApplied !== false && executionMayHaveApplied(error));
        response = failureResponse(request.id, error, {
          maybeApplied,
          ...taskDelta,
          events: boundedEvents(events),
          imports: [...imports.splice(0)],
          artifacts: [...artifacts],
          evidence: [...evidence],
          audit: {
            sequence: currentSequence,
            sourceSha256: sha256(request.code),
            maybeApplied,
            ...(interruptedRequest ? { interruptedRequest } : {}),
          },
        });
      }
      await persistTerminal({ request, response, sequence: currentSequence });
      if (interruptedRequest) interruptedRequest = null;
      return response;
    },
    async close() {
      if (closed) return;
      closed = true;
      try {
        if (currentSequence === 0) {
          await writeCheckpoint({
            protocol: REPL_PROTOCOL_VERSION,
            sessionId,
            sequence: 0,
            workspaceRoot,
            taskRoot,
            sessionRoot: sessionRootPath,
            state: snapshotState(state),
            artifacts,
            evidence,
            taskId: task.manifest.id,
            parentSessionId: task.parentSessionId,
          });
        }
      } finally {
        deregisterHooks();
        await lock.release();
      }
    },
  };

  async function appendJournal(record) {
    const line = `${JSON.stringify(record)}\n`;
    await assertJournalBudget(journalPath, parsed.maxJournalBytes, Buffer.byteLength(line));
    await appendFile(journalPath, line, { mode: 0o600 });
    if (process.platform !== "win32") await chmod(journalPath, 0o600);
  }

  async function persistTerminal({ request, response, sequence: requestSequence }) {
    const snapshot = {
      protocol: REPL_PROTOCOL_VERSION,
      sessionId,
      sequence: requestSequence,
      workspaceRoot,
      taskRoot,
      sessionRoot: sessionRootPath,
      state: snapshotState(state),
      artifacts: [...artifacts],
      evidence: [...evidence],
      last: {
        id: request?.id ?? null,
        sourceSha256: typeof request?.code === "string" ? sha256(request.code) : null,
        source: typeof request?.code === "string" ? request.code : null,
        response,
      },
      taskId: task.manifest.id,
      parentSessionId: task.parentSessionId,
      updatedAt: new Date().toISOString(),
    };
    await writeCheckpoint(snapshot);
    maybeInterruptForTest("checkpoint-after-rename");
    await appendJournal({
      protocol: REPL_PROTOCOL_VERSION,
      type: "request.terminal",
      sessionId,
      sequence: requestSequence,
      id: request?.id ?? null,
      sourceSha256: typeof request?.code === "string" ? sha256(request.code) : null,
      ok: response.ok,
      maybeApplied: response.error?.maybeApplied ?? false,
      at: new Date().toISOString(),
    });
  }

  async function writeCheckpoint(snapshot) {
    await atomicWriteJson(checkpointPath, snapshot);
  }

  return session;
}

async function executeCell(code, ctx, scopedConsole) {
  const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
  const cell = new AsyncFunction("ctx", "console", code);
  return cell(ctx, scopedConsole);
}

function parseReplArguments(args) {
  const options = {
    workspaceRoot: undefined,
    codeFile: undefined,
    taskId: undefined,
    newTaskGoal: undefined,
    maxRequestBytes: DEFAULT_MAX_REQUEST_BYTES,
    maxResponseBytes: DEFAULT_MAX_RESPONSE_BYTES,
    maxJournalBytes: DEFAULT_MAX_JOURNAL_BYTES,
    help: false,
  };
  const values = [...args];
  while (values.length > 0) {
    const value = values.shift();
    if (value === "--help" || value === "-h") options.help = true;
    else if (value === "--workspace") options.workspaceRoot = requiredOption(values, value);
    else if (value.startsWith("--workspace=")) options.workspaceRoot = value.slice(12);
    else if (value === "--file") options.codeFile = requiredOption(values, value);
    else if (value.startsWith("--file=")) options.codeFile = value.slice(7);
    else if (value === "--new") options.newTaskGoal = requiredOption(values, value);
    else if (value.startsWith("--new=")) options.newTaskGoal = value.slice(6);
    else if (value === "--max-request-bytes") options.maxRequestBytes = parsePositiveLimit(requiredOption(values, value), value);
    else if (value.startsWith("--max-request-bytes=")) options.maxRequestBytes = parsePositiveLimit(value.slice(20), "--max-request-bytes");
    else if (value === "--max-response-bytes") options.maxResponseBytes = parsePositiveLimit(requiredOption(values, value), value);
    else if (value.startsWith("--max-response-bytes=")) options.maxResponseBytes = parsePositiveLimit(value.slice(21), "--max-response-bytes");
    else if (value === "--max-journal-bytes") options.maxJournalBytes = parsePositiveLimit(requiredOption(values, value), value);
    else if (value.startsWith("--max-journal-bytes=")) options.maxJournalBytes = parsePositiveLimit(value.slice(20), "--max-journal-bytes");
    else if (value.startsWith("-")) throw replError("invalid-option", `Unknown REPL option: ${value}.`);
    else if (options.taskId == null) options.taskId = value;
    else throw replError("invalid-argument", `Unexpected REPL argument: ${value}.`);
  }
  return options;
}

async function readCodeFileRequest(source, maximum) {
  const target = path.resolve(source);
  const descriptor = await lstat(target).catch((error) => {
    if (error?.code === "ENOENT") throw replError("invalid-code-file", `REPL code file does not exist: ${target}`);
    throw error;
  });
  if (descriptor.isSymbolicLink() || !descriptor.isFile()) {
    throw replError("invalid-code-file", `REPL code file must be a regular non-symlink file: ${target}`);
  }
  if (descriptor.size > maximum) throw replError("request-too-large", `REPL code file exceeds ${maximum} bytes.`);
  let code;
  try {
    code = new TextDecoder("utf-8", { fatal: true }).decode(await readFile(target));
  } catch (error) {
    throw replError("invalid-code-file", `REPL code file must contain valid UTF-8: ${error.message}`);
  }
  if (code.trim() === "") throw replError("invalid-code-file", "REPL code file must contain a non-empty JavaScript cell.");
  const request = JSON.stringify({ id: "file", code });
  if (Buffer.byteLength(request, "utf8") > maximum) throw replError("request-too-large", `REPL code-file request exceeds ${maximum} bytes.`);
  return request;
}

async function resolveSessionOptions(options) {
  const resolved = {
    maxRequestBytes: DEFAULT_MAX_REQUEST_BYTES,
    maxResponseBytes: DEFAULT_MAX_RESPONSE_BYTES,
    maxJournalBytes: DEFAULT_MAX_JOURNAL_BYTES,
    ...options,
  };
  if (resolved.help) return resolved;
  if (resolved.newTaskGoal != null && resolved.taskId != null) {
    throw replError("invalid-argument", "Choose --new <goal> or an existing task ID, not both.");
  }
  if (resolved.newTaskGoal == null && resolved.taskId == null) {
    throw replError("missing-task", "Choose an existing task ID or create one with --new <goal>. Run officekit tasks first.");
  }
  return resolved;
}

async function ensureDirectory(target, label) {
  const descriptor = await lstat(target).catch((error) => {
    if (error?.code === "ENOENT") return null;
    throw error;
  });
  if (descriptor?.isSymbolicLink() || (descriptor && !descriptor.isDirectory())) {
    throw replError("unsafe-path", `${label} must be a regular directory: ${target}`);
  }
  if (!descriptor) await mkdir(target, { recursive: true, mode: 0o700 });
  if (process.platform !== "win32") await chmod(target, 0o700);
  return realpath(target);
}

function assertContainedCanonical(candidate, root, label) {
  const relative = path.relative(root, candidate);
  if (
    relative === ".." ||
    relative.startsWith(`..${path.sep}`) ||
    path.isAbsolute(relative)
  ) {
    throw replError("unsafe-path", `${label} must remain inside its declared root.`);
  }
  return candidate;
}

function parseRequestLine(line, maximum) {
  if (Buffer.byteLength(line, "utf8") > maximum) {
    return { ok: false, id: null, raw: null, error: replError("request-too-large", `REPL request exceeds ${maximum} bytes.`) };
  }
  let raw;
  try {
    raw = JSON.parse(line);
  } catch (error) {
    return { ok: false, id: null, raw: null, error: replError("invalid-json", `REPL request is not valid JSON: ${error.message}`) };
  }
  const id = raw && typeof raw === "object" ? raw.id ?? null : null;
  if (raw == null || typeof raw !== "object" || Array.isArray(raw)) {
    return { ok: false, id, raw, error: replError("invalid-request", "REPL request must be a JSON object.") };
  }
  if (raw.protocol != null && raw.protocol !== REPL_PROTOCOL_VERSION) {
    return { ok: false, id, raw, error: replError("unsupported-protocol", `REPL protocol must be ${REPL_PROTOCOL_VERSION}.`) };
  }
  if (typeof raw.id !== "string" || raw.id.trim() === "" || raw.id.length > 128) {
    return { ok: false, id: null, raw, error: replError("invalid-request", "REPL request id must be a non-empty string of at most 128 characters.") };
  }
  if (typeof raw.code !== "string" || raw.code.trim() === "") {
    return { ok: false, id: raw.id, raw, error: replError("invalid-request", "REPL request code must be a non-empty string.") };
  }
  return { ok: true, id: raw.id, code: raw.code, raw };
}

function createScopedConsole(events) {
  const capture = (level, values) => {
    const text = values.map((value) => {
      try { return typeof value === "string" ? value : JSON.stringify(serializeValue(value)); }
      catch { return String(value); }
    }).join(" ");
    const bounded = text.length > DEFAULT_MAX_EVENT_BYTES
      ? `${text.slice(0, DEFAULT_MAX_EVENT_BYTES)}…`
      : text;
    events.push({ level, text: bounded, at: new Date().toISOString() });
  };
  return Object.freeze({
    log: (...values) => capture("log", values),
    info: (...values) => capture("info", values),
    warn: (...values) => capture("warn", values),
    error: (...values) => capture("error", values),
  });
}

function boundedEvents(events) {
  return events.slice(0, 128);
}

function successResponse(id, fields) {
  return { protocol: REPL_PROTOCOL_VERSION, id, ok: true, ...fields };
}

function taskResponseMarker(task) {
  return {
    commitId: task?.commit?.commitId ?? null,
    publication: task?.publication == null
      ? null
      : `${task.publication.commitId}\n${task.publication.artifactId}\n${task.publication.path}\n${task.publication.sha256}`,
  };
}

function taskResponseDelta(before, task) {
  const after = taskResponseMarker(task);
  return {
    ...(after.commitId !== before.commitId && task?.commit ? { commit: task.commit } : {}),
    ...(after.publication !== before.publication && task?.publication ? { publication: task.publication } : {}),
  };
}

function failureResponse(id, error, fields = {}) {
  return {
    protocol: REPL_PROTOCOL_VERSION,
    id,
    ok: false,
    error: serializeError(error, fields.maybeApplied ?? false),
    ...fields,
  };
}

function serializeError(error, maybeApplied = false) {
  return {
    code: typeof error?.code === "string" ? error.code : "execution-failed",
    message: error?.message || String(error),
    retryable: error?.retryable === true,
    maybeApplied,
    ...(error?.stack ? { stack: error.stack } : {}),
  };
}

function snapshotState(state) {
  const safe = Object.create(null);
  const omitted = [];
  for (const [key, value] of Object.entries(state)) {
    if (isJsonSafe(value)) safe[key] = value;
    else omitted.push(key);
  }
  return { safe, omitted };
}

function isJsonSafe(value, seen = new WeakSet()) {
  if (value == null || typeof value === "string" || typeof value === "boolean") return true;
  if (typeof value === "number") return Number.isFinite(value);
  if (typeof value !== "object") return false;
  if (seen.has(value)) return false;
  seen.add(value);
  if (Array.isArray(value)) return value.every((entry) => isJsonSafe(entry, seen));
  if (!isPlainObject(value)) return false;
  return Object.values(value).every((entry) => isJsonSafe(entry, seen));
}

function serializeValue(value, depth = 0, seen = new WeakSet()) {
  if (value == null || typeof value === "string" || typeof value === "boolean") return value;
  if (typeof value === "number") return Number.isFinite(value) ? value : { type: "non-finite-number" };
  if (typeof value === "bigint") return { type: "bigint", value: value.toString() };
  if (typeof value === "function") return { type: "function", name: value.name || null, restorable: false };
  if (depth > 8) return { type: "depth-limit" };
  if (seen.has(value)) return { type: "circular" };
  seen.add(value);
  if (isFileBlobLike(value)) {
    return {
      type: "file-blob",
      mime: typeof value.type === "string" ? value.type : "application/octet-stream",
      bytes: value.bytes?.byteLength ?? null,
      metadata: isPlainObject(value.metadata) ? serializeValue(value.metadata, depth + 1, seen) : {},
      published: false,
    };
  }
  if (value instanceof Error) return serializeError(value);
  if (Array.isArray(value)) return value.slice(0, 512).map((entry) => serializeValue(entry, depth + 1, seen));
  if (!isPlainObject(value)) {
    return { type: "non-serializable", class: value.constructor?.name || "Object", restorable: false };
  }
  const result = {};
  for (const key of Object.keys(value).slice(0, 512)) result[key] = serializeValue(value[key], depth + 1, seen);
  return result;
}

function isFileBlobLike(value) {
  return value && typeof value === "object" && value.bytes instanceof Uint8Array && typeof value.arrayBuffer === "function";
}

function isPlainObject(value) {
  if (value == null || typeof value !== "object") return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

async function publishArtifact(value, options, { roots, sequence, artifacts }) {
  const destination = resolveOutputPath(options.path, roots.outputRoot, sequence, options.name);
  const sourcePaths = Array.isArray(options.sourcePaths) ? options.sourcePaths : [];
  const bytes = await readPublishBytes(value);
  await assertSafeOutput(destination, roots.outputRoot, [...sourcePaths, typeof value === "string" ? value : null]);
  await writeAtomic(destination, bytes, 0o600);
  const descriptor = {
    path: destination,
    kind: normalizeArtifactKind(options.kind, value),
    mime: options.mime || value?.type || "application/octet-stream",
    bytes: bytes.byteLength,
    sha256: sha256(bytes),
    locator: options.locator ?? null,
    visualReview: options.visualReview ?? "unavailable",
    reviewVerdict: options.reviewVerdict ?? null,
    contentView: options.contentView ?? "not-requested",
    limitations: Array.isArray(options.limitations) ? [...options.limitations] : [],
  };
  artifacts.push(descriptor);
  return descriptor;
}

async function readPublishBytes(value) {
  if (typeof value === "string") {
    const descriptor = await lstat(value);
    if (descriptor.isSymbolicLink() || !descriptor.isFile()) throw replError("unsafe-source", "Published source must be a regular file.");
    return readFile(value);
  }
  if (value instanceof Uint8Array) return value;
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (value && typeof value.arrayBuffer === "function") return new Uint8Array(await value.arrayBuffer());
  throw replError("invalid-artifact", "ctx.publish accepts a FileBlob, byte array, ArrayBuffer, or regular file path.");
}

function resolveOutputPath(requested, outputRoot, sequence, name) {
  const candidate = requested == null
    ? path.join(outputRoot, name || `artifact-${sequence}.bin`)
    : path.resolve(outputRoot, requested);
  const relative = path.relative(outputRoot, candidate);
  if (relative === ".." || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) {
    throw replError("unsafe-output", "Published artifact path escapes outputRoot.");
  }
  return candidate;
}

async function assertSafeOutput(destination, outputRoot, sourcePaths) {
  const relative = path.relative(outputRoot, destination);
  if (relative === ".." || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) throw replError("unsafe-output", "Published artifact path escapes outputRoot.");
  const existing = await lstat(destination).catch((error) => error?.code === "ENOENT" ? null : Promise.reject(error));
  if (existing?.isSymbolicLink()) throw replError("unsafe-output", "Refusing to publish through a symbolic link.");
  if (existing) throw replError("output-exists", `Published artifact already exists: ${destination}`);
  const canonicalParents = await realpath(path.dirname(destination));
  if (canonicalParents !== outputRoot && !canonicalParents.startsWith(`${outputRoot}${path.sep}`)) throw replError("unsafe-output", "Published artifact parent escapes outputRoot.");
  for (const source of sourcePaths.filter(Boolean)) {
    const sourcePath = path.resolve(source);
    const sourceCanonical = await realpath(sourcePath).catch(() => sourcePath);
    const destinationCanonical = path.join(canonicalParents, path.basename(destination));
    if (sourceCanonical === destinationCanonical) throw replError("source-overwrite", "Published artifact must not overwrite an input.");
  }
}

async function recordEvidenceFile(target, metadata, { evidenceRoot, evidence }) {
  if (typeof target !== "string" || target.length === 0) {
    throw replError("invalid-evidence", "Evidence path must be a non-empty relative path.");
  }
  const locator = validateEvidenceLocator(metadata?.locator);
  const visualReview = metadata?.visualReview ?? "unavailable";
  if (!["complete", "unavailable", "requires-human"].includes(visualReview)) {
    throw replError("invalid-evidence", "visualReview must be complete, unavailable, or requires-human.");
  }
  const candidate = path.resolve(evidenceRoot, target);
  const relative = path.relative(evidenceRoot, candidate);
  if (relative === ".." || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) throw replError("unsafe-evidence", "Evidence path escapes evidenceRoot.");
  const descriptor = await lstat(candidate).catch((error) => {
    if (error?.code === "ENOENT") return null;
    throw error;
  });
  if (!descriptor || descriptor.isSymbolicLink() || !descriptor.isFile()) throw replError("invalid-evidence", "Evidence must be an existing regular file.");
  const canonical = await realpath(candidate);
  if (canonical !== evidenceRoot && !canonical.startsWith(`${evidenceRoot}${path.sep}`)) throw replError("unsafe-evidence", "Evidence path escapes evidenceRoot.");
  const bytes = await readFile(canonical);
  const record = {
    path: canonical,
    kind: metadata?.kind || "evidence",
    locator,
    visualReview,
    bytes: bytes.byteLength,
    sha256: sha256(bytes),
  };
  evidence.push(record);
  return record;
}

function validateEvidenceLocator(value) {
  if (value == null) return null;
  if (!isPlainObject(value)) throw replError("invalid-evidence", "Evidence locator must be a plain object.");
  const allowed = new Set(["page", "slide", "sheet", "range", "cell", "object"]);
  for (const key of Object.keys(value)) {
    if (!allowed.has(key)) throw replError("invalid-evidence", `Evidence locator key is not supported: ${key}`);
  }
  for (const key of ["page", "slide"]) {
    if (value[key] !== undefined && (!Number.isSafeInteger(value[key]) || value[key] < 1)) {
      throw replError("invalid-evidence", `Evidence locator ${key} must be a positive integer.`);
    }
  }
  for (const key of ["sheet", "range", "cell", "object"]) {
    if (value[key] !== undefined && (typeof value[key] !== "string" || value[key].length > 512)) {
      throw replError("invalid-evidence", `Evidence locator ${key} must be a bounded string.`);
    }
  }
  return structuredClone(value);
}

async function atomicWriteJson(target, value) {
  await ensureDirectory(path.dirname(target), "checkpoint parent");
  const existing = await lstat(target).catch((error) => error?.code === "ENOENT" ? null : Promise.reject(error));
  if (existing?.isSymbolicLink() || (existing && !existing.isFile())) {
    throw replError("unsafe-path", `Checkpoint must be a regular file: ${target}`);
  }
  const temporary = path.join(path.dirname(target), `.${path.basename(target)}.${randomUUID()}.tmp`);
  await writeFile(temporary, `${JSON.stringify(value, null, 2)}\n`, { encoding: "utf8", mode: 0o600 });
  maybeInterruptForTest("checkpoint-before-rename");
  try {
    await rename(temporary, target);
  } finally {
    await rm(temporary, { force: true });
  }
}

function maybeInterruptForTest(point) {
  if (process.env.NODE_ENV === "test" && process.env.OFFICE_KIT_REPL_TEST_INTERRUPT_AT === point) {
    // Deliberately bypass finally blocks to model a process termination at the
    // exact point where a real host could be interrupted. The test-only guard
    // keeps this fault injector out of normal CLI behavior.
    process.exit(86);
  }
}

async function writeAtomic(target, bytes, mode = 0o600) {
  await ensureDirectory(path.dirname(target), "artifact parent");
  const existing = await lstat(target).catch((error) => error?.code === "ENOENT" ? null : Promise.reject(error));
  if (existing?.isSymbolicLink() || existing) {
    throw replError("output-exists", `Published artifact already exists: ${target}`);
  }
  const temporary = path.join(path.dirname(target), `.${path.basename(target)}.${randomUUID()}.tmp`);
  await writeFile(temporary, bytes, { mode });
  try {
    await rename(temporary, target);
  } finally {
    await rm(temporary, { force: true });
  }
}

async function assertJournalBudget(target, maximum, additionalBytes = 0) {
  const descriptor = await lstat(target).catch((error) => error?.code === "ENOENT" ? null : Promise.reject(error));
  if (descriptor?.isSymbolicLink() || (descriptor && !descriptor.isFile())) {
    throw replError("unsafe-path", "REPL session journal must be a regular file.");
  }
  if ((descriptor?.size ?? 0) + additionalBytes > maximum) {
    throw replError("checkpoint-too-large", `REPL checkpoint journal exceeds ${maximum} bytes.`);
  }
}

function executionMayHaveApplied(error) {
  return !new Set([
    "invalid-json", "invalid-request", "unsupported-protocol", "invalid-option",
    "invalid-argument", "remote-import", "unsafe-import", "unpublished-subpath",
    "invalid-artifact", "unsafe-output", "output-exists", "source-overwrite",
    "unsafe-source", "invalid-evidence", "unsafe-evidence", "invalid-input",
    "unsafe-input", "input-too-large", "source-changed", "artifact-exists",
    "invalid-artifact-id", "invalid-artifact-kind", "artifact-kind-mismatch",
    "invalid-review", "review-too-large", "stale-review", "review-failed", "unreviewed-artifact",
    "stale-commit", "invalid-commit", "revision-corrupt",
  ]).has(error?.code);
}

function createLazyExcelFacade() {
  let facade;
  let loading;
  const load = async () => {
    if (facade) return facade;
    loading ??= import("../excel-live/repl.mjs").then(({ createExcelLiveReplFacade }) => {
      facade = createExcelLiveReplFacade();
      return facade;
    });
    return loading;
  };
  return Object.freeze({
    doctor: (...args) => load().then((value) => value.doctor(...args)),
    sessions: (...args) => load().then((value) => value.sessions(...args)),
    execute: (...args) => load().then((value) => value.execute(...args)),
    disconnect: (...args) => load().then((value) => value.disconnect(...args)),
  });
}

function createLazyPowerPointFacade() {
  let facade;
  let loading;
  const load = async () => {
    if (facade) return facade;
    loading ??= import("../powerpoint-live/repl.mjs").then(({ createPowerPointLiveReplFacade }) => {
      facade = createPowerPointLiveReplFacade();
      return facade;
    });
    return loading;
  };
  return Object.freeze({
    doctor: (...args) => load().then((value) => value.doctor(...args)),
    sessions: (...args) => load().then((value) => value.sessions(...args)),
    execute: (...args) => load().then((value) => value.execute(...args)),
    disconnect: (...args) => load().then((value) => value.disconnect(...args)),
  });
}

function requiredOption(values, option) {
  const value = values.shift();
  if (value == null || value.startsWith("-")) throw replError("invalid-option", `${option} requires a value.`);
  return value;
}

function parsePositiveLimit(value, option) {
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed <= 0) throw replError("invalid-option", `${option} requires a positive integer.`);
  return parsed;
}

async function writeJsonLine(output, value, maximum) {
  const line = JSON.stringify(value);
  if (Buffer.byteLength(line, "utf8") > maximum) {
    const fallback = JSON.stringify(failureResponse(value?.id ?? null, replError("response-too-large", `REPL response exceeds ${maximum} bytes.`), { maybeApplied: Boolean(value?.error?.maybeApplied) }));
    output.write(`${fallback}\n`);
    return;
  }
  if (!output.write(`${line}\n`)) await once(output, "drain");
}

function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

function normalizeArtifactKind(kind, value) {
  if (["document", "workbook", "presentation", "pdf"].includes(kind)) return kind;
  const candidate = value?.metadata?.artifactKind;
  if (["document", "workbook", "presentation", "pdf"].includes(candidate)) return candidate;
  const mime = value?.type || "";
  if (mime.includes("word") || mime.includes("document")) return "document";
  if (mime.includes("sheet") || mime.includes("excel")) return "workbook";
  if (mime.includes("presentation") || mime.includes("powerpoint")) return "presentation";
  if (mime === "application/pdf") return "pdf";
  return "unknown";
}

function replError(code, message, options = {}) {
  const error = new Error(message);
  error.code = code;
  if (options.retryable != null) error.retryable = options.retryable;
  if (options.maybeApplied != null) error.maybeApplied = options.maybeApplied;
  return error;
}
