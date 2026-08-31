# Presentation Template Library

This plugin contains thirty-one original OfficeKit presentation style Skills.
Each template is guidance plus visual calibration evidence. A template may
also declare one reviewed clean-room reference PPJ/PPTX when native reuse and
provenance justify the package cost. `officekit template search --kind
presentation` discovers the styles in place; `officekit init` does not copy
them into each project.

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

Thirty styles were independently rebuilt from high-level observations of a
user-supplied visual reference set. Each now includes an OfficeKit-authored
clean-room reference PPTX made from unrelated content. The external reference
archive, source descriptions, page images, names, and geometry are not
distributed. Every shipped guide, calibration page, and reference deck is an
original OfficeKit work.

Evidence Ledger is the first PPJ-native reference template. Its source program,
compiled deck, and figures are OfficeKit-original and explicitly synthetic.
