# OfficeKit REPL

Use one `officekit repl` process for a multi-step artifact task. It keeps the
OfficeKit object graph and task helpers alive while keeping the transport
machine-readable. The process reads JSONL from standard input and writes one
JSON response per request to standard output; diagnostics belong in the
response `events` or on standard error.

```bash
officekit repl --workspace /absolute/project --task-root /absolute/task
```

Each line has a non-empty `id` and JavaScript `code`. `code` runs with
top-level `await`, `ctx`, and a scoped `console`:

```json
{"id":"inspect","code":"const {DocumentFile} = await ctx.import('office-kit'); ctx.state.doc = await DocumentFile.importDocx(await ctx.import('node:fs/promises').then(fs => fs.readFile(ctx.inputRoot + '/input.docx'))); return ctx.state.doc.inspect();"}
```

Use `return` for the value reported as `result`. Keep reusable helpers and
live OfficeKit objects in `ctx.state`; lexical variables from an earlier line
are not part of the persistence contract. A failed cell does not end the
process. Its response reports `retryable` and `maybeApplied`; reread the
affected artifact or range before retrying a mutation.

## Context and publication

`ctx` contains `sessionId`, `workspaceRoot`, `taskRoot`, `inputRoot`,
`assetRoot`, `outputRoot`, `evidenceRoot`, and `checkpointRoot`.

- `ctx.import(specifier)` permits published `office-kit` exports, `node:`
  built-ins, and local modules/dependencies inside the workspace. URLs,
  absolute paths, traversal, and unpublished OfficeKit subpaths are rejected.
- `ctx.publish(fileBlobOrPath, options)` writes an atomic, distinct output
  under `outputRoot` and returns an absolute path, type, byte count, SHA-256,
  locator, and `visualReview` status. It never overwrites an input.
- `ctx.recordEvidence(path, metadata)` registers an existing regular file
  under `evidenceRoot` with a SHA-256 and optional page/slide/sheet/range
  locator. Evidence is not a final deliverable.

Use the selected domain workflow for inspect → edit → re-read/verify → render
and then publish the final artifact:

```js
const { DocumentFile } = await ctx.import("office-kit");
ctx.state.doc ??= await DocumentFile.importDocx(`${ctx.inputRoot}/input.docx`);
const summary = await ctx.state.doc.inspect();
// Make a bounded edit only after inspecting the target.
const output = await ctx.state.doc.exportDocx();
return await ctx.publish(output, {
  name: "output.docx",
  kind: "document",
  visualReview: "unavailable",
});
```

The Excel Live Control facade is available only when a cell explicitly uses
`ctx.excel`. It exposes typed `doctor()`, `sessions()`, `execute(request)`,
and `disconnect(sessionId)` methods over the existing protocol. It does not
evaluate arbitrary Office.js. Run `officekit excel install` or
`officekit excel uninstall` separately; starting a REPL never installs a
certificate, trusts a root, starts a bridge, or downloads a provider.

## Checkpoints and resume

After every request, OfficeKit atomically updates a private checkpoint under
`checkpointRoot`. `checkpoint.json` contains the logical session, sequence,
source hash/text, safe JSON state, artifact/evidence references, and the last
response. `session.jsonl` records request starts and terminal audit records.

```bash
officekit repl --resume /absolute/task/.officekit-repl/<session-id>/checkpoint.json
```

Resume restores JSON-safe state and references without replaying source. Live
OfficeKit objects, functions, streams, and Excel sessions are process-local;
reconstruct them explicitly and inspect before another mutation. An unmatched
request-start record is surfaced as `maybeApplied: true` on the next response.

Template search, provider setup, rendering helpers, and Excel installation are
explicit commands. They are not hidden inside a REPL cell.
