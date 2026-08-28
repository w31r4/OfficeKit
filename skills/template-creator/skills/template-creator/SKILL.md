---
name: template-creator
description: Create or update a reusable local Office artifact template. For presentations, create a clean-room visual grammar from a style guide and representative PNG examples; for Word and Excel, retain a validated reference file. Use when the user asks to make or update an artifact-template skill. Do not use for one-off artifact creation from an existing template.
---

# Template Creator

Create or update a reusable local template. Presentation templates and document/spreadsheet
templates have deliberately different evidence models:

- A presentation template is a clean-room visual grammar: `SKILL.md`,
  `artifact-template.json` schema v3, `agents/agent.yaml`, `assets/preview.png`,
  and hashed `assets/examples/*.png`. It does not publish the source PPTX,
  MJS, DSL, fixed layout, or cloneable page.
- A DOCX/XLSX template remains reference-backed and uses schema v2. A PPTX may
  still be processed as a one-off reference for import/edit work, but publishing
  it as a new reusable presentation template uses the clean-room path below.

The uploaded source and all analysis evidence stay in the task workspace; only
the selected presentation guidance and original calibration images leave the task.

## Run template creation in one task

For preview, analysis, metadata, and packaging steps that span several commands,
use `officekit repl` and `../office-kit/references/repl.md`. Keep source paths,
style evidence, preview hashes, and selection metadata in the task state; do not
publish a template until its assets and selection card have been checked. Use
`ctx.recordEvidence` for preview or validation reports and return the created
Skill directory as an absolute path. The deterministic helper command remains
explicit; a REPL cell does not silently download or install a provider.

Use `workspaceRoot`, `taskRoot`, `inputRoot`, and `assetRoot` from
`../office-kit/references/workspace.md`. Keep the uploaded reference read-only, put
previews and temporary renders under `taskRoot`, and report the created Skill
directory as an absolute path.

## Routing

- Manage only direct-child template skills below `${OFFICE_KIT_HOME:-~/.office-kit}/skills`.
- Create a new template by default. Use a numbered name instead of overwriting an existing template.
- Update only when the user explicitly identifies exactly one existing `artifact-template-*` skill.
- Keep template creation local. Do not fetch remote templates or modify installed caches.

## Create workflow

### Presentation clean-room workflow

Use this path when the reusable result should teach an Agent a visual language
rather than ship a source deck. The style guide must be written in English and
state evidenced rules for hierarchy, surfaces, geometry, density, imagery, and
anti-patterns. Produce four to six original calibration pages from unrelated
content, render them to PNG, and inspect the spread. Example images are visual
evidence only; they are not pages to copy.

Run:

```bash
officekit run "$SKILL_DIR/scripts/create-template-skill.mjs" \
  --kind presentation \
  --style-guide-path "/absolute/path/style-guide.md" \
  --examples-json '[{"path":"/absolute/path/example-01.png","role":"opening claim"}]' \
  --preview-path "/absolute/path/preview.png" \
  --display-name "Evidence Ledger" \
  --description "Create evidence-led decision presentations with readable data and clear conclusions."
```

The command writes schema v3 with content-addressed example records and no
reference path. Verify the generated `SKILL.md`, `artifact-template.json`,
`agents/agent.yaml`, preview, and example PNGs. Keep `editProfile.level` at
`copy-only`; a clean-room style guide does not prove imported edit capability.

### Reference-backed Office workflow (DOCX/XLSX only)

1. Require exactly one `.docx` or `.xlsx` reference unless the user explicitly requests a batch. For a batch, complete this workflow separately for every file. An extension alone is not evidence: the creator must accept the reference as a bounded Office OPC package before retaining it. A PPTX may be inspected as one-off source material for the clean-room presentation path, but it is never retained in a reusable presentation template.
2. Infer a concise display name, intended-use description, and artifact kind
   from the reference and request. Always prepare schema-v2 selection metadata.
   Write `useWhen`, `avoidWhen`, audiences, content shapes, tone, and structure
   as concise English search text, regardless of the user's language. Include
   at least one evidence-backed `useWhen`; keep unknown arrays empty and
   density or color mode as `mixed`.
3. Create `preview.png` before packaging:
   - DOCX: render the reference and use a representative page PNG.
   - XLSX: render the used range of the first visible non-empty sheet.
4. Inspect the PNG. Stop if it is blank, clipped, corrupted, or not representative of the reference.
5. Set `SKILL_DIR` to this skill directory and pass shell-escaped values directly to the creator:

```bash
officekit run "$SKILL_DIR/scripts/create-template-skill.mjs" \
  --reference-path "/absolute/path/reference.docx" \
  --preview-path "/absolute/path/preview.png" \
  --display-name "Standup" \
  --description "Run a structured daily standup with updates, blockers, and owners."
```

Pass the complete selection metadata as one shell-escaped JSON value:

```bash
officekit run "$SKILL_DIR/scripts/create-template-skill.mjs" \
  --reference-path "/absolute/path/reference.docx" \
  --preview-path "/absolute/path/preview.png" \
  --display-name "Quarterly Review" \
  --description "Review quarterly performance, decisions, risks, and outlook." \
  --selection-json '{"useWhen":["quarterly business review"],"avoidWhen":["project kickoff"],"audiences":["executive"],"contentShapes":["KPIs","decisions","risks"],"visualTraits":{"tone":["formal"],"density":"medium","colorMode":"light","structure":["sectioned"]},"visualCommitment":"neutral","editProfile":{"level":"copy-only","verifiedOperations":[]},"provenance":{"license":"user-provided","source":"local-user-reference"}}'
```

Do not invent metadata just to fill a field. Do not claim a verified edit
operation from visual inspection. Keep
`editProfile.level` as `copy-only` until a real import/edit/export/reimport test
proves a narrower or broader profile. The script permits a minimal default only
when the English intended-use description itself is enough for `useWhen`; the
normal Skill workflow should pass the explicit evidence-backed profile.

6. Read the JSON result. For DOCX/XLSX, verify that the generated directory contains `SKILL.md`, schema-v2 `artifact-template.json`, `agents/agent.yaml`, the retained `assets/reference.<ext>`, and `assets/preview.png`; verify the recorded reference and preview hashes. For a clean-room presentation, verify schema v3, `assets/examples/*.png`, and the example/preview hashes instead.

The creator verifies ZIP CRCs, the required family-specific primary part and
content type, and exactly one root `officeDocument` relationship before it
acquires its write lock. A renamed text file, cross-family package, corrupted
archive, or broken primary relationship must fail closed without creating or
changing a template.

## Update workflow

Schema-v3 presentation templates are intentionally immutable style packages for
their first release. To revise one, create a new calibration set and invoke the
clean-room create path with a new display name or explicit skill name; do not
replace its examples in place without rerendering and rechecking the style
guide. Reference-backed DOCX/XLSX updates continue below.

1. Resolve the exact passed DOCX/XLSX template and read its `SKILL.md`, `artifact-template.json`, `agents/agent.yaml`, retained reference, and preview. Stop if it is not a direct child of the local skills directory or if more than one target was passed. Presentation schema-v3 packages are revised by creating a new calibration set through the clean-room path, not by replacing their examples in place.
2. Preserve the template folder name and every file or behavior the user did not ask to change.
3. For reference or visual changes, edit a temporary copy of the retained reference using the matching Office artifact workflow, render a new preview, and inspect it. For display-name or intended-use changes, retain the existing reference and preview unless they also change.
4. Pass every current or changed required value to the creator explicitly.
   Existing schema-v2 selection metadata is preserved when `--selection-json`
   is omitted; pass a complete English replacement value when that metadata
   must change:

```bash
officekit run "$SKILL_DIR/scripts/create-template-skill.mjs" \
  --mode "update" \
  --skill-name "artifact-template-standup" \
  --reference-path "/absolute/path/updated-reference.docx" \
  --preview-path "/absolute/path/updated-preview.png" \
  --display-name "Standup" \
  --description "Run a structured daily standup with updates, blockers, and owners."
```

5. The script accepts schema-v1 templates for migration, validates the existing template kind, preserves additional template-owned files, emits schema v2, and replaces the template atomically without changing its skill name.
6. Verify every requested change and confirm that no staging or backup directories remain.

## Response

Report the created or updated template's display name, artifact kind, absolute
path, and reference/preview hashes. State that the reference and preview remain
with the template, and briefly describe how to invoke the returned template
Skill. Do not emit host-specific cards, links, or sharing directives.

## Constraints

- Do not create an intermediary request file; pass creator inputs through command-line flags, including optional selection JSON.
- Do not delete or sanitize a retained DOCX/XLSX reference; fidelity depends on retaining it verbatim. A PPTX used as clean-room source material is not a retained reference.
- Do not change the artifact kind during an update.
- Do not mark a template `bounded-edit` or `composable` without repeatable capability evidence.
- Do not create translated metadata copies or a `searchLanguage` field.
- Do not modify global skill metadata or protocol files.
