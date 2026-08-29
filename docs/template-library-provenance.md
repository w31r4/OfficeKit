# Template library provenance

## Source-backed DOCX and XLSX templates

`skills/default-template-library/` retains 13 templates from
[`office-artifact-tool` `256cb31bfe0a07b3cef0051b6b159342be381378`](https://github.com/w31r4/office-artifact-tool/commit/256cb31bfe0a07b3cef0051b6b159342be381378):
seven DOCX and six XLSX. The pinned
[`reference/office-artifact-tool`](../reference/office-artifact-tool) submodule
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

`skills/presentation-template-library/` contains thirty-nine OfficeKit-original,
AGPL-licensed presentation Template Skills. Presentation schema v3 has one
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
A template may additionally ship one declared, reviewed clean-room PPJ and the
PPTX compiled from it. It never ships a user/source deck, MJS, executable page
code, SVG page skeleton, fixed Layout, or undeclared cloneable component.
Existing IDs are retained only as catalog identity. The guidance and
calibration pages were rewritten and rendered as OfficeKit-original work rather
than copied from the removed source-backed PPTX templates.

`presentation-template-creator` packages the same fixed surface from a distilled
guide and four to six original calibration images. Source references, analysis,
temporary artifacts, and review evidence remain task-local. Only an explicitly
licensed clean-room `referenceProgram`/`referencePptx` pair crosses that boundary.

## Verification

`test/default-template-library.mjs` checks the 13 source-backed DOCX/XLSX
templates, retained hashes, materialization, and bounded native workflows.
`test/template-creator.mjs` checks deterministic presentation schema-v3
creation/update and generic PPTX routing. `test/office-kit-skill.mjs` verifies
schema-specific discovery, hash validation, old presentation schema rejection,
and zero-or-one selection. Package checks require all thirty-nine presentation
styles and 197 calibration PNGs; they reject undeclared PPTX/PPJ, executable
code, and SVG page skeletons from template directories.
