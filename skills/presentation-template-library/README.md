# Presentation Template Library

This plugin contains 32 original OfficeKit presentation style Skills. Each
template is guidance plus visual calibration evidence, not a retained deck or
layout program. `officekit template search --kind presentation` discovers the
styles in place; `officekit init` does not copy them into each project.

The catalog is authored as clean-room OfficeKit material. It preserves broad,
public design vocabulary while avoiding copied prose, palettes, page geometry,
code, source decks, screenshots, or assets from third-party products.

Every nested template has exactly this surface:

```text
SKILL.md
artifact-template.json
agents/agent.yaml
assets/preview.png
assets/examples/*.png
```

The selected style informs a new deck-specific Design Grammar. The
Presentations Skill still composes every page from the current content and
reviews the rendered result. Selecting no template remains valid.
