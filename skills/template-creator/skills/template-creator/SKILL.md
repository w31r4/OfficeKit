---
name: template-creator
description: Create or update a reusable local document or spreadsheet template from one DOCX or XLSX reference. Use when the user asks to retain a Word or Excel reference for later reuse. Route every PPTX or presentation-style request to the installed presentation-template-creator Skill.
---

# Template Creator

Create source-backed DOCX and XLSX templates. Keep the verified Office reference
and representative PNG byte-for-byte so future work can import or clone it
faithfully. PowerPoint uses a different template form and must be routed to
`presentation-template-creator`.

## Routing

- Accept `.docx` and `.xlsx` only.
- For `.pptx`, presentation screenshots, presentation style descriptions, or
  OfficeKit presentation tasks, load `../presentation-template-creator/SKILL.md`.
  In repository plugin layout use
  `../../../presentation-template-creator/skills/presentation-template-creator/SKILL.md`.
- Manage only direct-child template Skills below
  `${OFFICE_KIT_HOME:-~/.office-kit}/skills`.
- Create a new numbered name by default. Update only when the user explicitly
  identifies exactly one existing template.

## Create

1. Keep the source file read-only and verify it as a bounded DOCX or XLSX OPC
   package. An extension alone is not evidence.
2. Render one representative PNG and inspect it for clipping, corruption, or a
   misleading crop.
3. Write one English canonical search profile: `useWhen`, `avoidWhen`, audience,
   content shape, tone, structure, density, and color mode. Unknown evidence
   remains empty or `mixed`.
4. Run:

```bash
officekit run "$SKILL_DIR/scripts/create-template-skill.mjs" -- \
  --reference-path "/absolute/path/reference.docx" \
  --preview-path "/absolute/path/preview.png" \
  --display-name "Standup" \
  --description "Run a structured daily standup with updates, blockers, and owners." \
  --selection-json '<complete-schema-v2-selection-json>'
```

5. Verify `SKILL.md`, schema-v2 `artifact-template.json`,
   `agents/agent.yaml`, `assets/reference.<ext>`, and `assets/preview.png`,
   including both recorded hashes.

The script validates ZIP CRCs, the family-specific primary part, content type,
and exactly one root `officeDocument` relationship before acquiring its write
lock. A renamed text file, cross-family package, corrupted archive, or broken
primary relationship fails without changing a template.

## Update

Resolve the exact local template, preserve its kind and unrelated owned files,
edit a temporary source copy through the matching Documents or Spreadsheets
workflow, render a new preview when visuals change, then rerun with:

```bash
officekit run "$SKILL_DIR/scripts/create-template-skill.mjs" -- \
  --mode update \
  --skill-name "artifact-template-standup" \
  --reference-path "/absolute/path/updated-reference.docx" \
  --preview-path "/absolute/path/updated-preview.png" \
  --display-name "Standup" \
  --description "Run a structured daily standup with updates, blockers, and owners."
```

Existing schema-v2 selection metadata is preserved when selection JSON is
omitted. A changed profile must be supplied in full. Confirm no staging,
backup, or lock residue remains.

## Evidence and response

Report the display name, artifact kind, absolute Skill path, reference hash,
and preview hash. Do not claim `bounded-edit` or `composable` without a real
import/edit/export/reimport proof. Do not create translated metadata copies,
fetch remote templates, or change global protocol files.
