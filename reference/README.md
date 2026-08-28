# Reference materials

This directory contains project-internal reference material for building `office-kit`.

## `office-artifact-tool` submodule

`reference/office-artifact-tool` is a Git submodule pointing at the public MIT-licensed `office-artifact-tool` reference package:

- Remote: `https://github.com/w31r4/office-artifact-tool.git`
- Purpose: behavior/API/workflow reference for creating a publishable open-source clean-room replacement.

Use this submodule to observe the reference package's public package shape, exported API surface, smoke tests, examples, and observable behavior.

The currently pinned reference revision is
`73c99c67ca7bbaa82cec0b158c647db583dcd970`, the remotely reachable
`origin/main` commit **Sync Office artifact runtime to 2.8.36**. Its package
manifest is `office-artifact-tool@2.8.36`; it keeps the reference-native
Documents, Spreadsheets, Presentations, PDF, Template Creator, and Default
Template Library trees in sync with the current public runtime payload.

The MIT-licensed, repository-only Default Template Library introduced at
`256cb31` remains available in this checkout as a reference source, but its
seven PPTX entries are not part of the current OfficeKit package. The package
retains only the 13 source-backed DOCX/XLSX templates under
`skills/default-template-library`; presentation styles live in the separate
AGPL-licensed `skills/presentation-template-library` and contain only clean-room
guides and PNG calibration evidence. The old PPTX references, executable Grid
layouts, and their materialization paths are retired from the published tree.
Pinning the exact remote commit keeps the current reference checkout
reproducible without importing the reference runtime or redistributing its
presentation payloads.

Do **not** vendor the reference package's runtime artifact, runtime module, runtime bindings, or implementation details into `office-kit`. Implement behavior independently using public standards, public libraries, OOXML/PDF specs, OpenXML SDK, Microsoft Office native automation, Playwright, LibreOffice, Poppler, PDF.js, sharp/canvas, and other legally usable technologies.

## Reference Skill source

The pinned submodule is the sole upstream reference Skill source. Its commit and
complete Skill-tree hashes are recorded in `skills/reference-sync.json` and
verified by `scripts/reference-skill-sync.mjs`. Project-adapted runnable Skills
live under `skills/`; PromptBench copies the pinned upstream Skill directly for
its reference subject, then patches only the package name inside the isolated
trial. Historical handoff snapshots are retained in Git history rather than as
a second live Skill tree.
