# OfficeKit task REPL

Use a durable OfficeKit task for multi-step artifact work. A task belongs to
one workspace and survives a new Agent context; one REPL process is only the
current session inside that task.

## Find or start the work

For a request to continue, revise, or finish prior Office work, inspect the
current workspace before creating anything:

```bash
officekit tasks --json
```

The default result contains at most five recent tasks. Match by the user's
goal, named input, and source SHA. If several tasks remain plausible, ask the
user. Do not choose a global latest task: tasks are visible only inside the
selected workspace.

Open a matching task or explicitly create a new one:

```bash
officekit repl t_7f2c9a31b804 --file phase-2.mjs
officekit repl --new "Create a promotion defense presentation" --file phase-1.mjs
```

With `--file`, write one ordinary UTF-8 JavaScript cell with top-level `await`;
do not hand-escape that code into JSONL. The command emits `session.ready`,
executes the file as one cell, emits its terminal response, and exits. Read the
ready record's task ID, brief, current
publishable commit descriptor, restored artifact paths, pending failures,
constraints, prior operation records, and next action before sending code. A
missing task ID fails; it never creates a new task by typo.

Use `officekit run task.mjs` instead for a genuinely one-shot script that does
not need a durable editing context.

Without `--file`, `officekit repl` is a long-lived JSONL process: a bare
invocation waits for JSONL on standard input and is not a one-shot command. Use
that form only when several cells must share `ctx.state` in one process. For a
quick probe, pipe one or more cells explicitly:

```bash
printf '%s\n' '{"id":"probe","code":"return 1;"}' | officekit repl --new "Probe"
```

For a resumed task, use a new `--file` invocation and reimport the restored
reviewed revision. Use `officekit run` when no durable task is required.

## Execute cells

Both `--file` and JSONL cells support top-level `await`, `ctx`, and a scoped
`console`. The low-level stream reads `{id, code}` JSON objects from standard
input and writes one JSON response per cell.

Use `ctx.state` for process-local functions and live OfficeKit objects reused
by later cells in the same process. A new session does not restore the JavaScript heap or
replay prior code. It restores reviewed file revisions; explicitly import those
files again.

`ctx` contains the task brief, workspace paths, lazy domain imports, and the
minimal artifact workflow:

- `ctx.task`: current compact task state and latest reviewed head.
- `ctx.import(specifier)`: published OfficeKit exports, `node:` built-ins, and
  local workspace modules; URLs, traversal, and private package paths fail.
- `ctx.input(path, options)`: copy a regular source into the task as an
  immutable input and return its artifact ID, path, type, size, and SHA-256.
  It does not retrieve an artifact already committed by this task.
- `ctx.commit(candidate, options)`: promote a candidate only when its OfficeKit
  review is non-failing and its delivery SHA matches the candidate bytes.
  Every commit requires a concise non-empty `options.summary`.
- `ctx.publish(commit, options)`: publish an artifact from the current reviewed
  task commit to a distinct final output.
- `ctx.recordEvidence(path, metadata)`: register bounded inspect, render, or QA
  evidence already written below `evidenceRoot`.

## Edit, review, commit, publish

For every meaningful edit batch:

```text
input or current committed revision
→ import and inspect
→ bounded typed edit
→ export a candidate
→ reopen and review
→ commit
```

Example shape:

```js
const { DocumentFile, reviewArtifact } = await ctx.import("office-kit");
const source = await ctx.input("/absolute/path/input.docx", {
  artifactId: "main-document",
});
ctx.state.document ??= await DocumentFile.importDocx(source.path);

// Inspect and apply one supported edit batch.
const candidate = await ctx.state.document.exportDocx();
const review = await reviewArtifact(candidate, {
  source: source.path,
  outputPath: `${ctx.taskRoot}/candidates/main-document.docx`,
  visualReview: "unavailable",
});
const commit = await ctx.commit(candidate, {
  artifactId: source.artifactId,
  summary: "Updated the requested section and preserved the remaining document",
  review,
  constraints: ["Keep the user template", "Do not change approved figures"],
  next: "Review the table on page 4",
});
ctx.state.commit = commit;
return commit;
```

Run semantic, structural, layout/render, optional content-reading, and
visual/human checks under the shared post-edit review contract. Request AnyDoc
only for a declared content-coverage gap. A failed or stale review cannot move
task HEAD; its candidate is reported as attention while the previous commit
remains the recovery point.

When the task is accepted, publish the current commit:

```js
return await ctx.publish(ctx.task.commit, { name: "final.docx" });
```

For a task commit containing several artifact heads, use
`{artifactId, name}` to publish each reviewed file. Publication never accepts
raw candidate bytes and never overwrites an input or existing output.

## Recovery and uncertainty

`officekit repl <task-id>` creates a new session and returns absolute paths for
the complete artifact snapshot in the latest reviewed commit. Pending failed
candidates are diagnostic inputs, not current revisions.

Inside a `--file` cell, reopen one reviewed artifact from the task state. Do
not pass its artifact ID to `ctx.input`:

```js
const path = await ctx.import("node:path");
const { FileBlob } = await ctx.import("office-kit");
const artifact = ctx.task.artifacts.find(({ id }) => id === "continued-deck");
const revision = artifact?.headRevision;
if (!revision) throw new Error("The reviewed continued-deck revision is missing");
const reviewedPath = path.resolve(ctx.taskRoot, revision.path);
const baseline = await FileBlob.load(reviewedPath);
```

After the final reimport and verification, publish with the current reviewed
descriptor, for example `ctx.publish(ctx.task.commit, { name: "final.pptx" })`.
An artifact ID or file path is not a publishable commit.

For an imported PPTX, reopen that reviewed revision and run `inspect` again
before every continued edit. Native leaf IDs and expected hashes are bound to
one revision; never reuse an ID from an earlier session. A successful lookup
after resume is the proof that the Agent rebuilt the node index from reviewed
bytes rather than relying on stale process state. Review and commit the new
candidate again before publication. `session.ready.operations` is immutable
audit evidence for prior Edit Plans, not a replacement for reinspection.

For PPTX, also read `slide.continuationCapability`: export/reimport a
`pending-clone`; use ready `bounded-overlay` only in a clean export. Commit with
a summary, reopen, and reinspect before other SlidePart edits.

Every cell still has a private atomic checkpoint and journal for crash
diagnosis. These are implementation details, not recovery keys. If a prior
request lacks a terminal record, the new task is `attention` and reports
`maybeApplied: true`; reimport and reread the affected artifact before another
mutation.

Template search, provider installation, Live Add-in setup, and rendering tools
remain explicit. Listing or opening a task does not initialize Office WASM,
MuPDF, providers, templates, or a Live bridge.
