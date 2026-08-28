# Default Template Library provenance

The default library contains one copy of the OfficeKit templates shipped with
the package. DOCX and XLSX entries remain reference-backed schema-v2 templates;
presentation entries are schema-v3 clean-room style packages.

## Presentation templates

The presentation entries are derived from public design guidance and observable
visual references from Kimi Slides, the public artifact-tool reference package,
and the project's independently licensed visual-reference research. The
published assets are OfficeKit-authored style guides and original calibration
images. They do not include the source decks, source code, MJS, DSL, fixed
layouts, or cloneable pages.

Each presentation entry contains:

- `SKILL.md` with the style grammar and anti-patterns;
- `artifact-template.json` schema v3 with English search metadata;
- `agents/agent.yaml`;
- `assets/preview.png` and five hashed `assets/examples/*.png` files.

The examples communicate hierarchy, surfaces, geometry, density, imagery, and
composition. They are visual evidence for the Agent, not files to copy. A
selected template is combined with the task's narrative and design grammar;
the Agent composes each new page with OfficeKit's native objects. A user brand
system or supplied reference takes precedence over a selected template, and
`none` is a valid selection when no style is suitable.

## Reference-backed templates

DOCX/XLSX entries retain their user-facing reference file and preview because
their workflow is defined by exact document or workbook structure. Their
reference hashes and preview hashes are recorded in
`skills/default-template-library/integrity.json`.

## Verification

`test/default-template-library.mjs` checks the canonical inventory, safe paths,
schema fields, PNG structure, every asset hash, and the aggregate hash.
Presentation entries are explicitly checked for the absence of PPTX, MJS, DSL,
and fixed-layout assets. `test/office-kit-skill.mjs` verifies hash-bound
discovery and that search returns candidates without choosing on the Agent's
behalf. `test/package-contents.mjs` checks the same inventory in the packed
package.

The former presentation reference decks and the fixed Grid Layout source tree
were retired from the live library. Their historical source and comparison
records remain available only in Git history or the read-only reference
submodule; they are not distributed as OfficeKit templates.
