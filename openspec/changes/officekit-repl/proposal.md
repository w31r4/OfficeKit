## Why

OfficeKit currently gives an Agent scripts and separate command surfaces, but it
does not provide one durable task context in which the Agent can inspect an
artifact, define a helper, apply several small edits, and verify the result.
That forces repeated process startup and re-imports, makes iterative repair
clumsy, and couples the workflow to whichever host happens to provide tools.

The project now has enough composable Office and PDF primitives to expose a
small, host-neutral JavaScript workbench. A task-level JSONL REPL gives every
supported Agent the same execution contract while keeping ordinary scripts and
explicit provider setup commands compatible.

## What Changes

- Add `officekit repl` as a local, JSONL-only JavaScript execution entry point
  for OfficeKit artifact tasks.
- Give each task a stable `ctx` workspace contract, explicit `ctx.state`, lazy
  package imports, safe artifact publication, evidence registration, and local
  session identifiers.
- Persist a private checkpoint after every request, including source text,
  JSON-safe state, artifact references, evidence, and error uncertainty; allow
  explicit checkpoint resume without replaying side effects.
- Continue processing after a code-cell failure and return structured
  `retryable` and `maybeApplied` information.
- Make `officekit run` a compatibility wrapper over the same package-resolution
  and execution policies rather than a competing Office workflow.
- Route DOCX, XLSX, PPTX, PDF, Template Creator, and Excel Live task workflows
  through the REPL; keep installation, template search, and provider setup as
  explicit control-plane commands.
- Add a lazy Excel Live facade that reuses the existing bridge protocol and
  never installs certificates or starts the bridge merely because a REPL starts.
- Add portable Skill guidance and examples for bootstrap, inspect/edit/verify
  loops, state reuse, publication, evidence, and failure recovery.

## Capabilities

### New Capabilities

- `officekit-repl`: Task-scoped JSONL JavaScript execution, workspace context,
  state reuse, result envelopes, checkpoints, explicit resume, and safe
  publication.
- `officekit-repl-excel-live`: Lazy, typed access to existing Excel Live
  sessions from a REPL without exposing raw Office.js or changing the bridge
  protocol.

### Modified Capabilities

- None. The repository has no existing OpenSpec capability specifications;
  `officekit run` compatibility and the seven Skill workflows are captured as
  requirements in the new capabilities above.

## Impact

- Adds a CLI protocol and session runtime under `src/cli/` and extends the
  package's documented execution surface.
- Changes the primary execution examples and workflow guidance in the seven
  distributed Skills, while preserving existing scripts and provider commands.
- Adds checkpoint files under task-local temporary directories and new REPL
  protocol, serialization, and integration tests.
- Reuses the current OfficeKit exports, PDF providers, Excel bridge, workspace
  contract, and `officekit run` resolver; it does not add a codec, change the
  Office wire protocol, or download runtimes.
