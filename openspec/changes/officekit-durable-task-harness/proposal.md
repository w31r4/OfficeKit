## Why

The current REPL can resume only when an Agent already knows an absolute private
checkpoint path, and it restores JSON-safe process state rather than a durable,
reviewed Office task. A fresh Agent context therefore cannot discover an
in-progress artifact job or safely continue from its last trustworthy file.

## What Changes

- Add project-local OfficeKit tasks under `.office-kit/tasks/`, discovered only
  within the selected workspace.
- Reduce the Agent-facing workflow to `tasks`, `repl`, `input`, `commit`, and
  `publish`.
- Add compact task listing, single-task detail, natural-language task creation,
  and opening an existing task from its latest reviewed commit.
- Persist immutable input snapshots, artifact revisions, review evidence,
  concise continuity notes, and per-process session lineage.
- Require a non-failing post-edit review before an artifact revision can become
  the task head or be published.
- Keep `officekit run` for one-shot work and keep live Office objects and helper
  functions process-local.
- **BREAKING**: remove path-oriented `officekit repl --resume` and the public
  `--task-root` workflow; callers open a task by ID instead. Do not add public
  discover, brief, history, latest, savepoint, or ephemeral commands.

## Capabilities

### New Capabilities

- `officekit-durable-task-harness`: Workspace-scoped task discovery, durable
  artifact revisions, reviewed commits, compact resume briefs, and safe
  publication through the minimal Agent harness.

### Modified Capabilities

<!-- The predecessor REPL change is complete but not yet synced as a main spec;
     this change supersedes its checkpoint/resume surface as one new contract. -->

## Impact

- Affects the OfficeKit CLI and REPL context, project-local `.office-kit`
  state, the OfficeKit coordination Skill, REPL examples, and focused package
  and portability tests.
- Does not change the Office wire protocol, OfficeKit Codec, PDF providers,
  template selection, Live request protocols, or file-format object models.
- Adds no runtime dependency and must remain lazy with respect to Office WASM,
  MuPDF, provider packs, and Live bridges.
