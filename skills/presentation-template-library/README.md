# Presentation Template Library

This plugin contains thirty-nine OfficeKit presentation style Skills: thirty
Kimi-derived style directions, eight Codex-aligned styles (seven source-bound
migrations and one OfficeKit clean-room reconstruction), and the OfficeKit-
original Evidence Ledger. Each template is guidance plus visual
calibration evidence. A template may also declare one reviewed reference
PPJ/PPTX when native continuation and provenance justify the package cost.
`officekit template search --kind presentation` discovers the styles in place;
`officekit init` does not copy them into each project. Reference PPJ/PPTX
packages remain in the GitHub source tree and are excluded from the npm
archive. Search is metadata-only and never downloads them; run
`officekit template fetch <template-id>` only when native continuation or a
reference render is needed.

Every nested template has exactly this surface:

```text
SKILL.md
artifact-template.json
agents/agent.yaml
assets/preview.png
assets/examples/*.png
assets/references/reference.ppj   # optional, declared and hash-bound
assets/references/reference.pptx  # optional, reviewed clean-room deck
```

The selected style informs a new deck-specific Design Grammar. The
Presentations Skill still composes every page from the current content and
reviews the rendered result. Selecting no template remains valid.

Thirty Kimi directions were independently rebuilt from high-level observations
of a user-supplied visual reference set. The seven Codex migrations retain
their MIT-licensed native source packages for source-bound inspection and
continuation; their current guides and calibration previews are packaged by
OfficeKit and do not silently claim a clean-room redraw. The external Kimi
reference archive and private analysis material are not distributed.

Evidence Ledger is the first PPJ-native reference template. Its source program,
compiled deck, and figures are OfficeKit-original and explicitly synthetic.
Optional references carry a `download` descriptor with a content hash and byte
count. The fetch command materializes the reference and every relative PPJ
source/asset dependency into an immutable local cache before compilation.
