# Default Template Library

This checked-in catalog provides 22 independently selectable Office templates: 7
documents, 9 presentations, and 6 spreadsheets. Most are reference-backed;
presentation templates created with schema v3 are clean-room visual grammars
with a style guide and representative example images, not retained source
decks.

Twenty templates are copied from the MIT-licensed `office-artifact-tool`
repository at commit `256cb31bfe0a07b3cef0051b6b159342be381378`
(`Add default Office template library`). Grid Layout Library is an additional
MIT presentation template generated from OfficeKit's existing 26-layout source
library and carries its own source commit in `artifact-template.json`. Every
retained or generated Office file and preview has an exact SHA-256 record in
`integrity.json`. See [LICENSE.md](LICENSE.md).

## Layout

```text
skills/default-template-library/
├── manifest.json
├── assets/icon.svg
└── skills/artifact-template-<name>/
    ├── SKILL.md
    ├── artifact-template.json
    ├── agents/agent.yaml
    └── assets/
        ├── preview.png
        ├── reference.docx | reference.pptx | reference.xlsx
        └── examples/*.png                 # schema-v3 presentations only
```

Each schema-v2 nested skill retains its reference Office file and preview image.
Its `artifact-template.json` adds English intended uses, avoid cases, audiences,
content shapes, visual traits, visual commitment, verified edit operations,
license/source provenance, and retained-asset hashes. A schema-v3 presentation
skill instead records a style guide, a preview, and one to eight hashed example
images; it deliberately contains no PPTX, MJS, DSL, fixed layout, or cloneable
page. OfficeKit can therefore shortlist templates without loading every Skill
description or opening every Office file.

Use the named template skill to create a new artifact while preserving the
retained layout and formatting unless the request calls for a change.

These resources ship once inside the global OfficeKit package.
`officekit init` leaves them there instead of copying them into every project.
Use `officekit template search` to query them, then create a distinct output
from the selected retained reference, or read the selected schema-v3 style guide
and examples before composing a new deck. Never overwrite or mutate a retained
reference or example asset.

For a guarded working copy, run:

```sh
officekit run skills/default-template-library/scripts/materialize-template.mjs \
  --template-id artifact-template-system-design \
  --output /absolute/path/system-design.docx
```

The materializer checks the retained source hash, refuses existing output and
audit paths, and writes a byte-identical working copy plus an audit record.
Use the matching Documents, Presentations, or Spreadsheets Skill to inspect,
edit, render, and verify that output. Complex source-bound Office graphs are
preserved only while unchanged; unsupported topology edits fail explicitly.
All seven imported PPTX templates expose at least one recognized SlidePart
placeholder whose existing text can be replaced through the bounded
Presentations workflow while its native identity, geometry, formatting, and
layout binding remain source-bound. Grid Layout Library instead exposes all 26
source-free layouts through the verified `source-slide-reuse` composable profile.
