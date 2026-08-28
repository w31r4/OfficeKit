# Default DOCX and XLSX Template Library

This checked-in catalog provides 13 independently selectable, source-backed
templates: seven documents and six spreadsheets. The retained Office files and
previews come from the MIT-licensed `office-artifact-tool` repository at commit
`256cb31bfe0a07b3cef0051b6b159342be381378`; exact asset hashes live in
`integrity.json`. See [LICENSE.md](LICENSE.md).

Presentation templates use the separate `presentation-template-library` and
its schema-v4 style-Skill format. This library intentionally contains no PPTX
templates.

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
        └── reference.docx | reference.xlsx
```

Each nested Skill keeps one source Office file and a schema-v2 search card.
OfficeKit can shortlist templates without opening every retained file.

The resources ship once inside OfficeKit. `officekit init` does not copy them
into each project. Search with `officekit template search`, then create a
distinct working output from the selected reference. Never overwrite the
reference.

For a guarded working copy, run:

```sh
officekit run skills/default-template-library/scripts/materialize-template.mjs \
  --template-id artifact-template-system-design \
  --output /absolute/path/system-design.docx
```

The materializer verifies the source hash, refuses existing output and audit
paths, and writes a byte-identical working copy plus an audit record. Use the
Documents or Spreadsheets Skill to inspect, edit, render, and verify it.
