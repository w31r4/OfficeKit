## ADDED Requirements

### Requirement: Workspace-scoped task discovery

OfficeKit SHALL discover durable tasks only below the selected workspace's
`.office-kit/tasks` directory. `officekit tasks` SHALL sort valid direct-child
tasks by descending update time, show at most five compact records by default,
support one-task detail and explicit `--all`, and never consult a global task
registry or initialize an Office/PDF/Live runtime.

#### Scenario: List recent tasks

- **WHEN** the current workspace contains eight valid tasks
- **THEN** `officekit tasks` reports the workspace, total count, five most
  recently updated compact rows, and that three more tasks exist

#### Scenario: Isolate another workspace

- **WHEN** another project contains its own `.office-kit/tasks`
- **THEN** listing from the current project does not reveal those tasks unless
  that project is explicitly selected as the workspace

#### Scenario: Inspect one task

- **WHEN** the Agent runs `officekit tasks <task-id> --json`
- **THEN** OfficeKit returns its goal, inputs, artifacts, latest reviewed
  commit, derived state, pending attention, next action, publication, and byte
  totals without returning cell source or document content

### Requirement: Minimal task-bound REPL lifecycle

`officekit repl --new "<goal>"` SHALL create a project-local task with an opaque
ID and start its first session. `officekit repl <task-id>` SHALL open an
existing task, acquire a single-writer lock, create a child session, and emit a
machine-readable resume brief before accepting cells. Missing tasks SHALL fail
instead of being created implicitly.

#### Scenario: Create a task

- **WHEN** the Agent starts a new REPL with a non-empty natural-language goal
- **THEN** OfficeKit creates one private task directory, assigns an opaque task
  ID, and reports an empty stable head in the startup envelope

#### Scenario: Resume in a fresh Agent context

- **WHEN** a new process opens an existing task that has a reviewed commit
- **THEN** the startup envelope returns the current publishable commit
  descriptor, restored artifact revisions, prior and current session lineage,
  pending candidates, constraints, and next action without requiring an old
  checkpoint path

#### Scenario: Reject concurrent writers

- **WHEN** a second REPL attempts to open a task whose writer lock is held
- **THEN** it fails before accepting code and does not modify task state

### Requirement: Immutable task inputs and revisions

`ctx.input(path)` SHALL stage a regular non-symlink input as a private,
read-only task copy, record its source path, size, type, and SHA-256, and return
a stable artifact ID. Revisions SHALL be immutable and identified by exact
content hashes; file hashes SHALL NOT determine workspace or task identity.

#### Scenario: Stage an input

- **WHEN** an Agent registers a PPTX source through `ctx.input`
- **THEN** OfficeKit copies its bytes into the selected task, returns an
  artifact descriptor, and later source changes cannot silently alter the
  staged input

#### Scenario: Reject an unsafe input

- **WHEN** the input is a directory, symlink, missing file, or exceeds the
  configured bound
- **THEN** registration fails without creating a task artifact revision

### Requirement: Reviewed artifact commits

`ctx.commit(candidate, options)` SHALL require an artifact ID, summary, review
report, and optional next action. It SHALL require a `passed` or
`passed-with-limitations` verdict whose delivery SHA-256 matches the candidate,
store one immutable revision, snapshot all artifact heads, and atomically
advance task HEAD. A failing or mismatched review SHALL preserve bounded
attention evidence and leave HEAD unchanged.

#### Scenario: Commit a reviewed edit

- **WHEN** an edited DOCX, XLSX, PPTX, or PDF has a matching non-failing review
- **THEN** OfficeKit records the revision and review, advances HEAD, updates the
  task summary and next action, and returns a commit descriptor

#### Scenario: Reject a failed review

- **WHEN** the candidate review verdict is `failed`
- **THEN** OfficeKit does not advance HEAD, reports the failed candidate as
  task attention, and the next session still restores the previous commit

#### Scenario: Reject a stale review

- **WHEN** the review delivery hash differs from the candidate bytes
- **THEN** OfficeKit rejects the commit as stale and does not write a stable
  revision

### Requirement: Resume from reviewed files

Opening an existing task SHALL restore descriptors for the complete artifact
head snapshot from its latest commit. It SHALL NOT claim to restore functions,
OfficeKit object graphs, streams, live host sessions, or automatically replay
prior code. Pending and interrupted candidates SHALL be reported but not made
current.

#### Scenario: Reconstruct a presentation

- **WHEN** a fresh process opens a presentation task at commit `c004`
- **THEN** the Agent receives the immutable PPTX revision path and explicitly
  reimports it before another edit

#### Scenario: Surface interrupted execution

- **WHEN** the prior session contains an unmatched request-start record
- **THEN** the task is reported as `attention` with `maybeApplied: true` and
  resumes from its last reviewed commit

### Requirement: Publish only reviewed commits

`ctx.publish` SHALL accept a commit descriptor for a task artifact and publish
that immutable revision to a distinct contained output path. It SHALL preserve
the commit review verdict, limitations, content-view state, and visual-review
state, and SHALL refuse candidates, foreign tasks, stale descriptors, source
paths, existing destinations, and path escapes.

#### Scenario: Publish a stable commit

- **WHEN** the Agent publishes the current reviewed PPTX commit
- **THEN** OfficeKit writes a distinct final file, returns its absolute path,
  type, size, SHA-256, and review state, and marks the task as published

#### Scenario: Reject an unreviewed candidate

- **WHEN** the Agent attempts to publish bytes that have not been committed
- **THEN** publication fails without changing the output directory or task
  publication state

### Requirement: One-shot compatibility and lazy startup

`officekit run` SHALL remain the one-shot workflow. Task listing, task detail,
new REPL startup, and resume startup SHALL not initialize MuPDF, OfficeKit Codec
WASM, provider packs, templates, or a Live bridge until task code explicitly
uses the corresponding capability.

#### Scenario: Run a one-shot script

- **WHEN** an existing task script runs through `officekit run`
- **THEN** package resolution, cwd, arguments, exit code, and error behavior
  remain compatible without creating a durable task

#### Scenario: Inspect tasks without runtimes

- **WHEN** the Agent runs `officekit tasks --json` in an initialized project
- **THEN** it reads only bounded local task metadata and does not initialize or
  download any artifact runtime
