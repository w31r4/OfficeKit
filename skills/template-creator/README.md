# Template Creator

Create or update reusable local Office templates. DOCX/XLSX templates retain a
validated reference file; presentation templates use a clean-room style guide,
preview, and representative PNG examples rather than publishing a source deck.

## Location and ownership

The creator writes user templates below:

```text
${OFFICE_KIT_HOME:-~/.office-kit}/skills/
```

Set `OFFICE_KIT_HOME` to choose another local root. Reference-backed templates
keep a verbatim copy of their Office reference and PNG preview. Clean-room
presentation templates keep only their style guide and original example images.
Choose an appropriate local storage location before creating one.

## Create

For a reference-backed DOCX/XLSX template, provide one supported Office
reference, a valid PNG preview, a concise display name, and an intended-use
description:

```sh
officekit run skills/template-creator/skills/template-creator/scripts/create-template-skill.mjs \
  --reference-path /absolute/path/reference.docx \
  --preview-path /absolute/path/preview.png \
  --display-name "Quarterly business review" \
  --description "Create a quarterly business review from the saved deck layout."
```

The command returns JSON containing the created template name, artifact kind, and local path. It selects a numbered name rather than overwriting an existing template.

Before it acquires a write lock or retains any bytes, the creator checks the
reference through the same bounded Office package inspector used by the public
facades. The extension must match a DOCX/XLSX OPC package with its required
primary part, declared main content type, and exactly one root
`officeDocument` relationship; ZIP entry CRCs are also verified. A renamed text
file, a cross-family Office package, a broken root relationship, or a corrupt
archive fails closed and creates no template tree. A PPTX used to study a
presentation style follows the clean-room path above and is never retained by
the published template.

The generated `artifact-template.json` uses schema version 2. Its searchable
fields use one English canonical form: `useWhen`, `avoidWhen`, audiences,
content shapes, tone, and structure. Unknown evidence stays empty or `mixed`;
there are no translated metadata copies. The sidecar also records SHA-256
values for the retained Office file and PNG. Without explicit selection
metadata, the English description becomes the sole `useWhen`; new templates
remain `copy-only` and visually `opinionated`.

Optional `--selection-json` accepts the complete selection profile in one JSON
value. Verified edit operations must come from a real
import/edit/export/reimport test, not visual inspection.

### Presentation clean-room create

First render unrelated calibration pages from the source material, write an
English style guide describing the evidenced visual grammar, and inspect the
PNG spread. Then create a presentation template without retaining the source
PPTX:

```sh
officekit run skills/template-creator/skills/template-creator/scripts/create-template-skill.mjs \
  --kind presentation \
  --style-guide-path /absolute/path/style-guide.md \
  --examples-json '[{"path":"/absolute/path/example-01.png","role":"opening claim"}]' \
  --preview-path /absolute/path/preview.png \
  --display-name "Evidence Ledger" \
  --description "Create evidence-led decision presentations with readable data and clear conclusions."
```

This emits schema version 3 with one to eight hashed example images. The
published template contains no PPTX, MJS, DSL, fixed layout, or cloneable page;
examples are visual evidence only. Keep `editProfile.level` as `copy-only`
unless a separate import/edit/export/reimport experiment proves otherwise.

## Update

Updates require the exact template name and preserve other skill-owned files:

```sh
officekit run skills/template-creator/skills/template-creator/scripts/create-template-skill.mjs \
  --mode update \
  --skill-name artifact-template-quarterly-business-review \
  --reference-path /absolute/path/updated-reference.docx \
  --preview-path /absolute/path/updated-preview.png \
  --display-name "Quarterly business review" \
  --description "Create a quarterly business review from the updated deck layout."
```

The creator validates artifact kind consistency, stages changes beside the
final directory, and replaces an updated template atomically with rollback if
placement fails. A per-home write lock prevents concurrent template writes.
Schema-v1 templates are migrated on update. Existing schema-v2 selection
metadata is preserved unless a complete `--selection-json` replacement is
provided.

Schema-v3 presentation templates are style packages rather than retained source
decks. To revise one, create a new calibration set and run the clean-room create
path again; do not replace its examples in place without rerendering and
rechecking the guide.

## Generated template layout

```text
$OFFICE_KIT_HOME/skills/artifact-template-<slug>/
├── SKILL.md
├── artifact-template.json
├── agents/agent.yaml
└── assets/
    ├── reference.docx | reference.xlsx
    ├── preview.png
    └── examples/*.png                 # schema-v3 presentations only
```

`artifact-template.json` records the supported kind, paths, selection evidence,
edit profile, provenance, and content hashes. The creator validates PNG chunk
structure and CRCs before copying images. Reference-backed templates also get
the fail-closed Office-package admission check on every create or update.
