## 1. Task store and discovery

- [x] 1.1 Add a bounded project-local task store with safe workspace discovery, private task directories, validated manifests, immutable files, and atomic writes.
- [x] 1.2 Add `officekit tasks`, compact five-row and JSON output, one-task detail, `--all`, explicit workspace selection, and guarded deletion.
- [x] 1.3 Add tests for workspace isolation, ordering and limits, invalid manifests, symlink/path escape rejection, deletion, and lazy runtime startup.

## 2. Task-bound REPL

- [x] 2.1 Replace Agent-facing path resume/task-root options with `repl --new <goal>` and `repl <task-id>` while retaining existing JSONL cell execution and safety limits.
- [x] 2.2 Add task writer locking, parent/child session lineage, interrupted request attention, and the versioned `session.ready` resume envelope.
- [x] 2.3 Add tests for creation, missing tasks, clean restart, concurrent writers, interruption, and no automatic code or object-graph replay.

## 3. Durable artifact workflow

- [x] 3.1 Implement `ctx.input` with immutable staged copies, stable artifact IDs, exact SHA-256 descriptors, bounds, and source protection.
- [x] 3.2 Implement `ctx.commit` with review/report validation, immutable deduplicated revisions, task-wide head snapshots, failed-candidate attention, and atomic HEAD advancement.
- [x] 3.3 Restrict task publication to current commit descriptors while preserving output containment, no-replace, hashes, limitations, and visual-review metadata.
- [x] 3.4 Add DOCX, XLSX, PPTX, PDF, multi-artifact, stale-review, failed-review, source-change, restart, and reviewed-publication tests.

## 4. Skill and package convergence

- [x] 4.1 Rewrite the OfficeKit REPL/workspace instructions around `tasks → repl → input → edit → review → commit → publish` and remove path-checkpoint terminology.
- [x] 4.2 Update focused CLI, Skill portability, package-contents, and reference-sync tests without changing domain codecs or shared Help/docs owned by concurrent work.
- [x] 4.3 Run OpenSpec validation, focused task/REPL/Skill gates, package smoke, and the complete `npm test`; record environment-only blockers without running unrelated slow provider release lanes.
