# Reference materials

This directory contains project-internal reference material for building `office-kit`.

## `office-artifact-tool` submodule

`reference/office-artifact-tool` is a Git submodule pointing at the public MIT-licensed `office-artifact-tool` reference package:

- Remote: `https://github.com/w31r4/office-artifact-tool.git`
- Purpose: behavior/API/workflow reference for creating a publishable open-source clean-room replacement.

Use this submodule to observe the reference package's public package shape, exported API surface, smoke tests, examples, and observable behavior.

The currently pinned reference revision is
`0f501d1a06d0566a90224a18b4c0d7a89671df66`, the remotely reachable
`origin/main` commit **Refresh Office skill guidance**. Its package manifest is
`office-artifact-tool@2.8.31`; it includes updated public PDF AcroForm
guidance, Presentation style/routing references, and Template Creator package
admission checks.

The MIT-licensed, repository-only Default Template Library introduced at
`256cb31` remains unchanged in this revision: 20 retained Office template Skills
containing 7 DOCX, 7 PPTX, and 6 XLSX references plus previews. This project
retains those assets byte-for-byte under `skills/default-template-library`,
records their original source hashes and license, and tests
import/export/edit/render behavior independently. The neutral
`grid-layout-library` naming and Template Creator workflow remain present and
adapted into this project's public Skills. Pinning the exact remote commit keeps
the current reference checkout reproducible without importing the reference
runtime.

Do **not** vendor the reference package's runtime artifact, runtime module, runtime bindings, or implementation details into `office-kit`. Implement behavior independently using public standards, public libraries, OOXML/PDF specs, OpenXML SDK, Microsoft Office native automation, Playwright, LibreOffice, Poppler, PDF.js, sharp/canvas, and other legally usable technologies.

## Reference Skill source

The pinned submodule is the sole upstream reference Skill source. Its commit and
complete Skill-tree hashes are recorded in `skills/reference-sync.json` and
verified by `scripts/reference-skill-sync.mjs`. Project-adapted runnable Skills
live under `skills/`; PromptBench copies the pinned upstream Skill directly for
its reference subject, then patches only the package name inside the isolated
trial. Historical handoff snapshots are retained in Git history rather than as
a second live Skill tree.
