## ADDED Requirements

### Requirement: JSONL task execution

`officekit repl` SHALL run one local task process that reads one JSON object per
input line and writes one JSON response per request line. Each request SHALL
contain a non-empty `id` and a string `code`; code SHALL support top-level
`await`, and requests SHALL execute in input order with no concurrent cells.

#### Scenario: Execute a code cell

- **WHEN** the Agent sends `{"id":"one","code":"return 1 + 1"}`
- **THEN** the REPL writes one response with protocol version `1`, the same
  request ID, `ok: true`, and a JSON result of `2`

#### Scenario: Reject malformed input

- **WHEN** a line is invalid JSON, lacks an ID, or has a non-string `code`
- **THEN** the REPL writes a structured validation error for that request and
  does not execute the supplied code

#### Scenario: Preserve the JSONL channel

- **WHEN** code calls `console.log` or a library emits diagnostic text
- **THEN** the REPL keeps stdout valid JSONL and places captured events in the
  response or diagnostics on stderr

### Requirement: Explicit lazy task context

Every request SHALL execute with a `ctx` containing the workspace roots,
`sessionId`, a state object, `import(specifier)`, artifact publication, and
evidence registration. OfficeKit, MuPDF, OpenChestnut, Excel bridge, and
provider modules SHALL NOT initialize before the code explicitly imports or
uses them.

`ctx.import` SHALL resolve published `office-kit` exports, permitted `node:`
modules, and local dependencies from the declared workspace. It SHALL reject
URLs, path traversal, and unpublished OfficeKit subpaths.

#### Scenario: Start without Office runtime initialization

- **WHEN** the Agent starts `officekit repl` and sends a state-only cell
- **THEN** the response succeeds without loading MuPDF, starting the Excel
  bridge, or initializing an Office codec runtime

#### Scenario: Import a published API on demand

- **WHEN** code executes `await ctx.import("office-kit")`
- **THEN** the REPL returns the installed package's published API and records
  the import in the request audit

#### Scenario: Reject an unpublished import

- **WHEN** code calls `ctx.import("office-kit/src/index.mjs")`
- **THEN** the REPL returns an unpublished-subpath error without resolving a
  private file

### Requirement: In-process state reuse

The REPL SHALL preserve `ctx.state` and in-memory OfficeKit objects across
successful and failed requests in the same task process. The documented stable
way to reuse a function or object SHALL be assigning it to `ctx.state`; lexical
variables from a previous cell SHALL NOT be a required contract.

#### Scenario: Reuse a helper function

- **WHEN** one cell assigns `ctx.state.makeSummary = rows => ...` and a later
  cell calls `ctx.state.makeSummary(rows)`
- **THEN** the later cell invokes the original function without redefining it

#### Scenario: Continue after an error

- **WHEN** a cell throws after possibly mutating an in-memory artifact
- **THEN** the REPL returns `ok: false` with `retryable` and `maybeApplied`,
  preserves the process state, and accepts the next request

### Requirement: Checkpoint and explicit resume

After every request, the REPL SHALL atomically write a private checkpoint that
contains the session ID, sequence, request ID, source text and hash,
JSON-safe state, result or error, artifact/evidence references, and audit
metadata. Non-JSON values SHALL remain process-local and SHALL NOT be presented
as restorable state.

`--resume` SHALL restore the latest JSON-safe state and references without
automatically replaying code that may have side effects. A resumed process
SHALL retain an auditable link to the prior checkpoint.

#### Scenario: Save a successful request

- **WHEN** a code cell completes successfully
- **THEN** the REPL publishes a new checkpoint before accepting the next cell
  and the checkpoint contains the cell source and result metadata

#### Scenario: Resume without duplicate mutation

- **WHEN** the Agent starts `officekit repl --resume <checkpoint>`
- **THEN** JSON-safe state and artifact references are restored, prior cells are
  not replayed, and non-serializable functions or Office objects are absent
  until the Agent explicitly reconstructs them

### Requirement: Safe artifact and evidence publication

`ctx.publish` SHALL write a final artifact to an absolute path, refuse to
overwrite an input or escape the declared output root, calculate SHA-256, and
return a stable artifact descriptor. `ctx.recordEvidence` SHALL only register
files under `evidenceRoot` and SHALL preserve locator and visual-review status.

#### Scenario: Publish a FileBlob

- **WHEN** code publishes a DOCX, XLSX, PPTX, or PDF `FileBlob`
- **THEN** the REPL writes it to a distinct output path and returns its kind,
  absolute path, size, and SHA-256

#### Scenario: Protect an input

- **WHEN** code asks `ctx.publish` to use an input path or a symlink alias of an
  input path
- **THEN** publication fails before writing and the input bytes remain unchanged

### Requirement: Compatibility with officekit run

`officekit run <task.mjs>` SHALL remain available for existing scripts and CI.
It SHALL reuse the REPL's published-package resolution, local dependency rules,
error classification, and audit conventions without requiring a project-local
`office-kit` installation.

#### Scenario: Run an existing task

- **WHEN** an existing `.mjs` task imports `office-kit` and a local dependency
  through `officekit run`
- **THEN** both imports resolve as before, the task receives its original argv
  and cwd, and no JSONL migration is required

#### Scenario: Reject an invalid task source

- **WHEN** `officekit run` receives a URL, stdin marker, private subpath, or
  non-JavaScript file
- **THEN** it returns the existing explicit validation error and does not fall
  back to another execution path
