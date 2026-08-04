## 1. CLI and session runtime

- [x] 1.1 Extract the published OfficeKit target map and local-dependency resolver from `src/cli/run-task.mjs` into a shared CLI module without changing existing `officekit run` behavior.
- [x] 1.2 Add `src/cli/repl.mjs` and dispatch `officekit repl` from `src/cli/officekit.mjs` with workspace, task-root, resume, and bounded-input options.
- [x] 1.3 Implement sequential JSONL parsing and protocol-v1 response envelopes with request IDs, top-level await, scoped console events, and stdout channel protection.
- [x] 1.4 Enforce request validation, URL/private-subpath rejection, code and response size limits, and deterministic error codes before executing a cell.
- [x] 1.5 Execute cells with explicit `ctx` and scoped console arguments; preserve only `ctx.state` and in-memory objects across requests.

## 2. Workspace, imports, and publication

- [x] 2.1 Build the REPL context from the existing workspace/evidence contract and generate a local UUID session ID without reading host thread or plugin identifiers.
- [x] 2.2 Implement `ctx.import` for published OfficeKit exports, permitted `node:` modules, and workspace-local dependencies with canonical path and traversal checks.
- [x] 2.3 Implement bounded result serialization for primitives, plain objects, errors, OfficeKit summaries, `FileBlob` values, circular values, and oversized values.
- [x] 2.4 Implement `ctx.publish` with output-root containment, source/input protection, atomic writes, SHA-256 descriptors, and artifact metadata.
- [x] 2.5 Implement `ctx.recordEvidence` with evidence-root containment, locator validation, and `visualReview` status recording.

## 3. Checkpoint and recovery

- [x] 3.1 Implement the private task checkpoint directory with append-only `session.jsonl` events and an atomically replaced `checkpoint.json` snapshot.
- [x] 3.2 Record request source, source hash, sequence, imports, result/error, artifact/evidence references, and `maybeApplied` status after every request.
- [x] 3.3 Implement explicit `--resume` that restores JSON-safe state and references, links the new process to the prior session, and never replays side-effecting code.
- [x] 3.4 Detect interrupted request-start records on resume and expose their uncertainty as `maybeApplied: true`.
- [x] 3.5 Add checkpoint permission, symlink, size, atomic-replacement, and interrupted-write protections for macOS, Linux, and Windows paths.

## 4. Excel Live integration

- [x] 4.1 Add `src/excel-live/repl.mjs` as a lazy facade over the current client, validator, bridge, and error modules.
- [x] 4.2 Expose typed `doctor`, `sessions`, `execute`, and `disconnect` operations through `ctx.excel`, preserving operation IDs, idempotency, audit fields, and `maybeApplied`.
- [x] 4.3 Ensure REPL startup and file-only cells do not create Excel state, trust certificates, start the bridge, upload manifests, or download anything.
- [x] 4.4 Return explicit unavailable-session and unsupported-platform errors without falling back to closed-file XLSX editing.

## 5. Compatibility and Skill migration

- [x] 5.1 Refactor `officekit run` to use the shared resolver and error classification while preserving argv, cwd, local dependency resolution, stack output, and private-subpath rejection.
- [x] 5.2 Add the canonical portable REPL reference and update OfficeKit, Documents, Spreadsheets, Presentations, PDF, Excel Live Control, and Template Creator guidance to use the REPL for artifact tasks.
- [x] 5.3 Update examples and workflow snippets to demonstrate `ctx.import`, `ctx.state`, inspect/edit/verify loops, `ctx.publish`, evidence registration, and uncertain retry handling.
- [x] 5.4 Keep template search, Excel installation/uninstallation, provider setup, and deterministic helper commands explicit and document their boundary with REPL execution.
- [x] 5.5 Add Skill portability checks preventing host-specific tools, thread IDs, remote imports, and claims that non-serializable state is resumable.

## 6. Tests and release metadata

- [x] 6.1 Add unit tests for JSONL parsing, top-level await, request ordering, console capture, validation errors, result serialization, and output-channel integrity.
- [x] 6.2 Add integration tests for state/helper reuse, FileBlob publication, source protection, evidence registration, checkpoints, explicit resume, and interrupted requests.
- [x] 6.3 Add clean-install tests proving REPL startup and state-only tasks do not initialize MuPDF, OpenChestnut, Excel bridge, or providers.
- [x] 6.4 Add mocked Excel bridge tests for lazy facade loading, session discovery, typed execution, disconnect, unavailable sessions, and error uncertainty.
- [x] 6.5 Extend CLI, Skill, package-contents, release, and deterministic-build gates for the new command, protocol docs, and checkpoint behavior.
- [x] 6.6 Bump package/CLI metadata to `0.6.0`, update English-first README/API/release documentation, and record the REPL protocol version without changing Office wire protocol.
- [x] 6.7 Run the complete npm, Skill, Office, PDF, Excel Live mock, package, release, and hosted CI gates before publishing the change.
