## Context

`officekit repl` currently creates a random process session, normally below a
system temporary task root, and resumes only from a caller-supplied
`checkpoint.json`. The checkpoint is useful for crash auditing but cannot be
discovered by a fresh Agent and cannot restore functions or OfficeKit object
graphs. OfficeKit already reserves project-local `.office-kit` state for Skill
ownership and PDF policy, so durable artifact tasks can use the same workspace
boundary without introducing a global registry.

## Goals / Non-Goals

**Goals:**

- Let a fresh Agent discover concise task summaries in the current project and
  open one task without knowing a session or checkpoint path.
- Restore immutable Office/PDF bytes from the latest reviewed commit, then let
  the Agent explicitly reimport the artifact.
- Keep the public harness small: `officekit tasks`, task-bound `officekit repl`,
  `ctx.input`, `ctx.commit`, and reviewed `ctx.publish`.
- Preserve lazy runtime loading, source immutability, deterministic error
  envelopes, interrupted-request evidence, and compatibility with
  `officekit run`.

**Non-Goals:**

- Restoring a JavaScript VM, function closures, streams, or live Office host
  sessions across processes.
- Global task search, shared or cloud task state, branching, merging, automatic
  code replay, or an adversarial JavaScript sandbox.
- Changes to format codecs, Office/PDF models, template selection, AnyDoc,
  providers, or Live protocols.

## Decisions

### Workspace directory is the namespace

Resolve an explicit `--workspace` first, otherwise walk from `cwd` to the
nearest `.office-kit` directory, then use the Git root when present, otherwise
`cwd`. Store tasks below `.office-kit/tasks/<task-id>`. Do not derive workspace
identity from file hashes and do not create a global recent-task index.

Each task contains a `task.json`, staged `inputs`, immutable `revisions`,
unpromoted `candidates`, `evidence`, and append-only `sessions`. Managed task
state is protected by a nested ignore file and private directory permissions.

### Task and session are separate identities

`officekit repl --new "<goal>"` creates an opaque task ID and the first process
session. `officekit repl <task-id>` opens the same task, validates its manifest
and HEAD, acquires a single-writer lock, creates a child session, and emits one
`session.ready` bootstrap envelope. Session IDs remain audit data and are never
the user-facing recovery key.

### Compact discovery replaces task-management verbs

`officekit tasks` scans direct child task manifests only, sorts by `updatedAt`,
and shows five compact rows. `officekit tasks <id>` shows the resume brief;
`--all` removes the display limit and `--json` emits the same bounded schema.
The command derives `new`, `stable`, `attention`, or `published` from stored
facts. There are no separate discover, brief, show, history, or latest verbs.

### Artifact bytes, not object graphs, are durable

`ctx.input(path)` copies one regular non-symlink source into the task, records
its original absolute path and SHA-256, and returns an artifact descriptor.
Each committed revision is an immutable file addressed by a SHA-256 recorded
in the task manifest. A commit updates one artifact head and snapshots all task
artifact heads, so a multi-artifact task resumes a consistent set.

`ctx.state` continues to reuse objects and helpers inside one process. Its
JSON-safe subset may remain in the private session journal for diagnostics but
is not the durable task contract and is not restored as authoritative state.

### Review is the commit boundary

`ctx.commit(value, { artifactId, summary, review, next })` accepts a path or
FileBlob candidate. It validates the review report, requires a non-failing
verdict and a delivery SHA matching the candidate, writes the immutable
revision, and atomically advances HEAD. A failed or mismatched review preserves
the candidate as attention evidence but leaves HEAD unchanged.

`ctx.publish(commitDescriptor, options)` publishes only a committed artifact
revision. It retains existing source-path, containment, no-replace, hash, and
visual-review protections. `passed-with-limitations` remains publishable only
with its limitations preserved; a task that explicitly requires aesthetic or
pixel confirmation cannot be called complete while visual review requires a
human.

### Old cell checkpoints become private audit records

Keep atomic per-cell session checkpoints and interrupted-write detection under
the selected task for crash diagnosis. Remove path-oriented `--resume` and
`--task-root` from the Agent-facing CLI rather than maintaining two recovery
models. No automatic migration is attempted for anonymous temporary sessions;
existing task scripts continue through `officekit run`.

## Risks / Trade-offs

- [Task directories can grow with large revisions] → Keep revisions immutable,
  deduplicate identical artifact bytes within a task, report byte totals in
  task detail, and require explicit `tasks --delete <id> --yes` cleanup.
- [A copied project also copies its local task state] → Use relative managed
  paths plus byte hashes, revalidate every external source path, and never use
  a global workspace identity.
- [Two Agents can corrupt one task] → Hold a private single-writer lock for the
  REPL lifetime and fail clearly when it is already held.
- [Agent-authored summaries can be inaccurate] → Treat goal, summary, and next
  as orientation only; derive revisions, hashes, Review status, timestamps,
  pending failures, and publication facts from OfficeKit records.
- [Arbitrary REPL code can read other workspace files] → Document task
  isolation as state organization, keep managed write APIs contained, and
  leave process sandboxing outside this change.

## Migration Plan

1. Add the task store and read-only `officekit tasks` command without loading
   file runtimes.
2. Bind new and existing REPL sessions to a task and emit the bootstrap brief.
3. Add input, commit, and reviewed publication primitives; update Skills and
   examples to the reduced flow.
4. Remove the documented path-resume/task-root surface and update package
   gates. The Office wire and Live request protocol versions do not change;
   the REPL transport version advances because of the startup envelope and
   task-bound request context.

Rollback removes the new task command and restores the previous REPL entry;
task directories remain ordinary local data and are not destructively deleted.

## Open Questions

None for V1. Global discovery, task sharing, branching, replay, and hard process
sandboxing are explicitly deferred.
