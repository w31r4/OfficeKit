# Template library provenance

## Source-backed DOCX and XLSX templates

`skills/default-template-library/` retains 13 templates from
[`office-artifact-tool` `256cb31bfe0a07b3cef0051b6b159342be381378`](https://github.com/w31r4/office-artifact-tool/commit/256cb31bfe0a07b3cef0051b6b159342be381378):
seven DOCX and six XLSX. The pinned
[`legacy/presentations/mjs/office-artifact-tool`](../legacy/presentations/mjs/office-artifact-tool) submodule
contains the authoritative source tree used by the byte-comparison gate.

The source repository declares MIT, Copyright (c) 2026 w31r4. The retained
license is at
[`skills/default-template-library/LICENSE.md`](../skills/default-template-library/LICENSE.md).
Each template preserves its source Office file and preview, while local
schema-v2 metadata adds evidence-backed search fields, exact hashes, provenance,
and verified operations. `integrity.json` records individual and aggregate
identity.

OfficeKit discovers these templates in place. Materialization verifies the
source hash, creates a distinct working file plus audit, and refuses overwrite.
Documents or Spreadsheets then owns the edit and review workflow.

## Original presentation Template Skills

`skills/presentation-template-library/` contains thirty-nine AGPL-licensed
presentation Template Skills. The set includes thirty Kimi-derived style
directions, eight Codex-aligned styles (seven source-bound migrations and one
OfficeKit clean-room reconstruction), and the OfficeKit-original Evidence
Ledger. Presentation schema v3 has one
public form:

```text
SKILL.md
artifact-template.json
agents/agent.yaml
assets/preview.png
assets/examples/*.png
assets/references/reference.ppj   # optional
assets/references/reference.pptx  # optional
```

The guide is the style authority; the images are visual calibration evidence.
The optional reference packages are stored in the GitHub repository and are
excluded from the npm archive. Their sidecar entries include a
`download.url`, byte count, and matching SHA-256. `officekit template search`
does not fetch them; `officekit template fetch <template-id>` downloads the
declared reference and every relative PPJ source/asset dependency into an
immutable local cache, verifying each hash before it becomes a compiler input.
Kimi directions and the Simple Dark pilot use OfficeKit-authored clean-room
calibration programs. The seven Codex migrations retain their declared
MIT source-bound PPJ/PPTX packages in the GitHub tree so native objects and
unsupported subgraphs remain inspectable and resumable; `template fetch`
materializes them only when needed. Their guides and previews are new OfficeKit
migration material, not a claim that the historical deck was redrawn. No
source-bound package is hidden behind a generic image-only preview.
The library never ships executable page code, SVG page skeletons, fixed Layout
instructions, or undeclared cloneable components.

`presentation-template-creator` packages the same fixed surface from a distilled
guide and four to six calibration images. Source references, analysis,
temporary artifacts, and review evidence remain task-local. A source-bound
`referenceProgram`/`referencePptx` pair crosses that boundary only when its
license, hash, and continuation value are declared in the sidecar.

## Verification

`test/default-template-library.mjs` checks the 13 source-backed DOCX/XLSX
templates, retained hashes, materialization, and bounded native workflows.
`test/template-creator.mjs` checks deterministic presentation schema-v3
creation/update and generic PPTX routing. `test/office-kit-skill.mjs` verifies
schema-specific discovery, lazy-reference descriptors, hash validation, old
presentation schema rejection, and zero-or-one selection. Package checks require
all thirty-nine presentation styles and every declared calibration PNG; reference
PPJ/PPTX trees are intentionally excluded from npm and undeclared decks,
executable code, and SVG page skeletons are rejected from template directories.
