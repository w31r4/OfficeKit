## Context

OfficeKit already has a local CLI, a published package exports map, a
workspace/evidence contract, the `officekit run` package resolver, and a
versioned Excel Live bridge. It does not yet have a process that can keep an
OfficeKit object graph and Agent-defined helpers alive across several requests.

The REPL must therefore add an execution layer without adding an Office object
model, a second codec, a daemon, or a host-specific integration. It must also
keep runtime loading lazy: starting a session is not permission to initialize
MuPDF, the Office WASM codec, the Excel bridge, or a managed provider.

## Goals / Non-Goals

**Goals:**

- Provide one task-scoped JSONL JavaScript execution protocol for artifact work.
- Make `ctx.state` the explicit in-process reuse contract and checkpoint only
  values that are safe to restore.
- Reuse the current package exports, workspace rules, `officekit run` resolver,
  Excel protocol, and artifact QA helpers.
- Give Skills a portable inspect/edit/verify/publish loop that works with any
  Agent capable of running a local command.
- Preserve explicit setup boundaries for Excel certificates and PDF providers.

**Non-Goals:**

- No background session server, remote transport, MCP server, or telemetry.
- No sandbox stronger than the existing `officekit run` local-code trust model.
- No automatic replay of side-effecting code during resume.
- No serialization of live OfficeKit objects, functions, streams, or Excel
  sessions.
- No raw `run_officejs`, provider installation, or silent XLSX fallback.

## Decisions

### 1. One sequential JSONL process

Add `src/cli/repl.mjs` and dispatch `repl` from `src/cli/officekit.mjs`.
`runReplCommand` owns one process, reads stdin with a bounded line reader, and
executes each request through a `ReplSession` in sequence. A request has
`{ protocol: 1, id, code }` where `protocol` may be omitted for compatibility;
responses always include it. Stdout is reserved for response lines. A scoped
console object captures `log`, `info`, `warn`, and `error` as bounded events;
uncaptured process diagnostics go to stderr.

Each cell is evaluated by an async function receiving `ctx` and the scoped
console. This gives top-level `await` and `return` while keeping cell-local
lexical variables out of the persistence contract. The Node process remains
the same, so functions and OfficeKit instances assigned to `ctx.state` stay
usable until EOF or process exit.

The alternative of a background daemon was rejected because it creates orphan
processes, a new authentication surface, and cross-task state leakage. A
separate worker per cell was rejected because it loses the object graph and
reintroduces import/export overhead.

### 2. Explicit context and package resolution

Extract the published-target mapping currently embedded in
`src/cli/run-task.mjs` into a shared resolver. `ctx.import(specifier)` uses it
for `office-kit` and its published subpaths, uses a workspace-anchored
`createRequire` for permitted local dependencies, and permits `node:` built-ins.
It rejects URLs, absolute paths, traversal, and unpublished package subpaths.

The resolver returns module namespaces through dynamic import. It never imports
the root package while constructing a session. The existing OfficeKit modules
remain responsible for their own deeper lazy imports, so a state-only cell does
not initialize MuPDF, the OfficeKit Codec, Excel, or a provider.

The alternative of an implicit `ctx.officekit` proxy was rejected because it
obscures import boundaries and makes runtime initialization harder to audit.

### 3. Workspace-scoped publication

Create a session context from the existing workspace contract. `ctx.publish`
accepts a `FileBlob` or an existing file plus an explicit or default output
name, resolves canonical paths, rejects input/source overlap and traversal,
writes atomically, and computes SHA-256. `ctx.recordEvidence` validates that a
file is under `evidenceRoot` and stores locator and visual-review metadata.

The serializer handles primitives, arrays, bounded plain objects, errors, and
known OfficeKit summaries. It represents a `FileBlob` by type, byte length, and
metadata unless it has been published. Circular, oversized, or live values are
reported as non-serializable rather than silently expanded.

### 4. Checkpoint as journal plus snapshot

Each task gets a private directory under `taskRoot`, containing:

- `session.jsonl`: append-only request-start, request-result/error, import,
  artifact, evidence, and audit events. Request source text is retained here.
- `checkpoint.json`: atomically replaced latest snapshot containing session
  metadata, JSON-safe `ctx.state`, artifact/evidence references, and the last
  completed sequence.

The response exposes the checkpoint directory path. Before executing a cell,
the session appends a start record. After success or failure it serializes the
safe state, replaces the snapshot through a temporary sibling and rename, then
appends the terminal record. On resume, an unmatched start record is treated as
`maybeApplied: true`; no code is replayed. Resume restores safe state and
references and assigns a new process instance to the same logical session.

The alternative of serializing arbitrary functions and Office objects was
rejected because it would claim fidelity that the underlying object graphs and
external runtimes cannot guarantee. Source text is retained for audit and
manual reconstruction, not automatic replay.

### 5. `officekit run` remains a compatibility adapter

Refactor `run-task.mjs` to consume the shared package-target resolver and error
classification. Existing task modules continue to receive their original
`process.argv`, cwd, local dependency resolution, and stack behavior. They do
not need to become JSONL files and do not silently route to another codec.

The new Skills use REPL cells for artifact workflows; deterministic helper
scripts, template search, Excel installation, and provider setup retain their
explicit commands. This keeps the execution core unified without pretending
that installation or a managed provider is an artifact code cell.

### 6. Excel Live is a lazy facade over the existing protocol

Add `src/excel-live/repl.mjs` with a factory that is dynamically imported when
`ctx.excel` is first used. It delegates `doctor`, `sessions`, `execute`, and
`disconnect` to the existing client, validator, and bridge request functions.
It preserves operation IDs, idempotency, typed failures, audit fields, and
`maybeApplied`; it never exposes arbitrary Office.js source.

`ctx.excel` may start or reuse the bridge only after an explicit Excel operation.
Certificate generation, trust, manifest upload, and uninstall remain under
`officekit excel install|uninstall`; no REPL startup or file task performs them.

### 7. Skill guidance is portable and task-oriented

Add a canonical REPL reference beside the OfficeKit Skill and add concise,
standalone bootstrap and recovery sections to the six domain/component Skills
that execute code. The guidance teaches:

1. start one REPL and keep its session ID;
2. import only the selected API with `ctx.import`;
3. store reusable helpers and live objects under `ctx.state`;
4. inspect before mutation and reread after uncertain mutation;
5. publish outputs and register evidence explicitly; and
6. resume by reconstructing non-serializable state instead of replaying it.

No Skill text names a particular host tool, thread, browser, or image provider.

## Risks / Trade-offs

- [Broad local code execution] → Document the same trust boundary as `run`,
  reject remote/unpublished imports, and keep control-plane setup explicit.
- [A failed cell may have mutated an artifact] → Always return
  `maybeApplied`, preserve state, and require inspect-before-retry guidance.
- [Checkpoint source can contain secrets] → Keep it task-local with private
  permissions, never upload it, and make the retention path visible in output.
- [Non-serializable state disappears on resume] → Expose this explicitly in the
  checkpoint and provide stable `ctx.publish` plus local helper-module imports.
- [JSONL output can be corrupted by library stdout] → Scope the console,
  reserve stdout, bound events, and test the channel with noisy fixtures.
- [Excel bridge availability varies by desktop host] → Preserve typed bridge
  errors and never fall back to closed-file editing for a live-session request.
- [Long sessions can grow journals] → Bound per-request source, result, event,
  and total journal sizes; return a clear limit error before accepting more.

## Migration Plan

1. Add the shared resolver, REPL session, protocol validator, serializer,
   checkpoint store, and lazy Excel facade behind the new CLI command.
2. Add focused REPL and Excel mock tests, then run the current `officekit run`
   and full package gates unchanged.
3. Update the seven Skill workflows and examples to use REPL cells for artifact
   work, retaining explicit setup/helper commands and compatibility examples.
4. Bump the package/CLI release to `0.6.0`, publish protocol documentation, and
   update package/release contents tests.
5. Rollback is removing the `repl` dispatch and Skill guidance; `run`, existing
   exports, Office wire protocol, and provider adapters remain independently
   usable throughout.

## Open Questions

No blocking design questions remain for the first implementation. A persistent
background session, cross-machine resume, and serializable Office object graphs
are explicitly deferred to a later change.
