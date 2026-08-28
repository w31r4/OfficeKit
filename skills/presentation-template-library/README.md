# Presentation Template Library

This plugin contains thirty original OfficeKit presentation style Skills. Each
template is guidance plus visual calibration evidence, not a retained deck or
layout program. `officekit template search --kind presentation` discovers the
styles in place; `officekit init` does not copy them into each project.

Every nested template has exactly this surface:

```text
SKILL.md
artifact-template.json
agents/agent.yaml
assets/reference.pptx
assets/preview.png
assets/examples/*.png
```

The selected style informs a new deck-specific Design Grammar. Each shipped
guide is self-contained: it records the communication territory, page
archetypes, visual carriers, layer order, density/rhythm, variation limits,
and review checks that make the style usable without reopening its source.
The examples calibrate range; they are not pages to copy. The Presentations
Skill still composes every page from the current content and reviews the
rendered result. Selecting no template remains valid.

Thirty styles were independently rebuilt from high-level observations of a
user-supplied visual reference set. Every style ships an OfficeKit-authored
native `reference.pptx` candidate whose layers can be inspected and edited.
The candidate evidence is kept honest: a style is not called restored merely
because it round-trips, and native-host or visual gaps remain explicitly
pending until they are proven. The reference archive, source descriptions,
page images, names, and geometry are not distributed. Each shipped guide,
calibration page, and reference deck is an original OfficeKit work.

Image-led styles use a role-aware pool of 19 original calibration photographs.
The batch author assigns distinct scenes to cover, evidence, visual, detail,
and closing roles before reusing an image; styles whose source explicitly avoids
photography remain native-vector or chart-led. This keeps the examples varied
without turning a picture count into a design rule.
