## Why

OfficeKit currently calls two incompatible things a PowerPoint template: a
retained source PPTX copied and edited through schema-v2 metadata, and a Grid
library made of fixed JavaScript layouts. That split constrains free authoring,
duplicates product concepts, enlarges the package, and keeps source assets that
OfficeKit does not need to redistribute.

## What Changes

- **BREAKING** Define one PowerPoint template form: an independently loadable
  Template Skill containing one `SKILL.md`, schema-v3 search metadata, and
  OfficeKit-original PNG examples. It contains no PPTX, layout code, DSL,
  source component, or cloneable page.
- Add a separately packaged `presentation-template-creator` Skill, installed by
  `officekit init`, that distills references and packages a validated style
  Skill without retaining its source deck.
- Keep the generic `template-creator` for DOCX and XLSX; route PPT requests to
  the specialist creator and reject presentation schema v2.
- Make template search return the selected Skill and its visual examples;
  Presentations derives a deck-specific Design Grammar and composes every page
  rather than materializing a reference file.
- Rebuild the seven bundled PPT templates and Grid as eight clean-room style
  Skills with the same IDs, then delete every bundled PPTX, old preview, fixed
  Grid module, registry, and fallback.
- Preserve reference decks, design systems, and source-bound continuation as
  distinct non-template workflows.

## Capabilities

### New Capabilities

- `presentation-template-skills`: The single schema-v3 PowerPoint template
  format, search result, selection semantics, and Presentations consumption
  workflow.
- `presentation-template-creator`: Reference distillation, original calibration
  examples, deterministic packaging, provenance, and update behavior for
  reusable PowerPoint style Skills.

### Modified Capabilities

None. The existing template behavior has not been archived into canonical
repository specs.

## Impact

- `officekit template search`, template metadata validation, default search
  roots, package contents, provenance records, and template-specific tests.
- A new default-installed Skill plugin plus OfficeKit/Presentations/template
  routing and Claude marketplace/reference-sync manifests.
- Removal of bundled presentation reference files and embedded Grid code while
  leaving DOCX/XLSX schema-v2 templates unchanged.
- Target release is `1.1.0`. There are no existing users, so the change carries
  no compatibility shim or deprecation period. Office wire protocol and public
  JavaScript Office APIs are unchanged.
