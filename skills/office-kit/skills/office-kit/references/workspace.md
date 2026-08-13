# OfficeKit workspace and evidence contract

Use these names in instructions, task files, examples, and handoffs. They are
ordinary local paths, not a host-specific API.

## Workspace roots

- `workspaceRoot`: the user's current directory, or an explicitly supplied
  project directory.
- `taskRoot`: the selected durable task below
  `workspaceRoot/.office-kit/tasks/<task-id>`; it contains private inputs,
  revisions, candidates, evidence, and session records.
- `inputRoot`: read-only source files and imported references.
- `assetRoot`: user-provided images, templates, fonts, and other reusable
  assets.
- `outputRoot`: final deliverables. Unless the user gives another destination,
  use `workspaceRoot/outputs`.
- `evidenceRoot`: renders, inspections, verification reports, and other QA
  evidence that is not itself a deliverable.
- `sessionId`: a local OfficeKit task identifier. Never derive it from a chat,
  thread, browser, or host identifier.

## Path rules

1. An explicit user path wins.
2. If no project is named, use `process.cwd()` as `workspaceRoot`.
3. Use `officekit tasks` to find resumable work and `officekit repl --new
   <goal>` to create a task. Use `officekit run` and `os.tmpdir()` only for
   intentionally one-shot work.
4. Treat inputs and retained templates as read-only. A final output must be a
   distinct path and must not overwrite an input.
5. Put renders and intermediate QA files under `evidenceRoot`; do not cite a
   temporary QA file as the final artifact.
6. Use absolute paths in scripts and result handoffs so a caller can relocate
   or link the artifact without guessing the working directory.
7. When resolving a relative path against a declared root, reject `..`
   traversal and symlink aliases that escape that root.

`SKILL_DIR` is allowed when a bundled script needs to locate its own files. It
is a Skill-relative path variable, not a workspace identity.

## Result and evidence envelope

Every completed artifact task returns, in the user's language, at least:

```json
{
  "artifact": {
    "path": "/absolute/path/to/output.docx",
    "kind": "document",
    "sha256": "..."
  },
  "evidence": [
    {
      "path": "/absolute/path/to/evidence/inspect.json",
      "kind": "inspect",
      "locator": { "page": 1 }
    }
  ],
  "reviewVerdict": "passed-with-limitations",
  "contentView": "anydoc",
  "visualReview": "complete"
}
```

`kind` is one of `document`, `workbook`, `presentation`, or `pdf`. A locator
may use `page`, `slide`, `sheet`, `range`, or another domain-specific address.
Use `visualReview: "unavailable"` when no visual input capability exists, and
`visualReview: "requires-human"` when the result needs human visual approval.
Never claim visual review from a structural report alone.

`reviewVerdict` is the machine review result from `reviewArtifact()` when that
facade is used. `contentView` is optional and may be `"anydoc"` only when the
Agent requested the lazy Markdown reading view. Neither field replaces the
owning Skill's domain checks or proves a visual review.

The host may turn an absolute artifact path into a link or citation. Skills do
not emit a host-specific message directive.
