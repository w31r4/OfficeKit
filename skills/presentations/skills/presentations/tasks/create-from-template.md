# Create with a template, design system, or reference deck

Use this route when one design authority or style source exists. Before drawing,
read the doctrine, shared visual floor, scenario policy, selected scenario guide,
and editorial trim Skill.

## 1. Classify the source

Choose exactly one primary source:

- `template`: one presentation Template Skill, its visual examples, and its
  OfficeKit-authored reference deck when declared;
- `design-system`: user or brand rules that are authoritative;
- `style-transfer`: a reference deck used only as visual evidence;
- `source-continuation`: an existing PPTX whose native pages/components remain
  the actual starting state.

Do not call all four “templates.” A design system overrides a conflicting style
Skill. Do not mix multiple Template Skills. Keep every uploaded file read-only
and task-bound by SHA-256.

## 2. Gather only the relevant evidence

For `template`, read its complete `SKILL.md`, preview, and four-to-six
role-labelled examples. The guide is the reusable design system: extract its
communication territory, hierarchy, palette roles, typography rhythm, geometry
and line behavior, density, visual carriers, imagery, chart language, motifs,
variation limits, and anti-patterns. Do not reduce it to a few adjectives or
copy a helper's default silhouette. If the template declares
`assets/reference.pptx`, import and inspect that OfficeKit-authored deck when
native backgrounds, source-derived assets, or editable layer order improve the
result. It is a starting source, not a fixed page recipe; do not clone every
page or treat it as an external source deck.

For `design-system`, record exact supplied rules and unresolved gaps. For
`style-transfer`, render and inspect the reference deck but do not copy its
wording, exact page geometry, or protected assets. For `source-continuation`,
import and inspect the package, then read
[reference-deck conditioned generation](../references/template-conditioned-generation.md)
and capability guidance before cloning or editing anything.

## 3. Write the current deck's plan

Use authoring mode `create-from-template`. Record the communication job,
scenario, selected source, evidence hashes, and one chosen direction. Write a
new Design Grammar for this deck:

- palette and surface roles;
- type hierarchy and rhythm;
- geometry and line rules;
- page-density rhythm and visual carriers;
- image, SVG, chart, diagram, and table treatment;
- allowed motifs and explicit anti-patterns.

If the source uses full-page imagery, scrims, crossing diagrams, or foreground
copy over evidence, inspect its cross-type layer order and read
[layered composition](../references/layered-composition.md). Treat those
relationships as part of the design grammar rather than flattening the page to
a screenshot.

Every page gets a claim, evidence, content budget, dominant carrier, and source
strategy. “Follow the template” is not a composition intent. Unknown source
facts remain unresolved.

## 4. Compose or continue

For `template`, `design-system`, and `style-transfer`, compose every page freely
with editable native objects. Use examples to understand relationships, not to
trace a calibration page. The selected style saves design reasoning; it never
pins page geometry.

For `source-continuation`, use inspected source-slide/component capabilities.
Export/reimport pending clones before another bounded edit. Preserve opaque
graphs and stop at an unsupported target instead of rebuilding the deck.

For decks longer than four pages, first render the opening, one evidence page,
and the densest/highest-risk page. Repair the same direction before expansion.

## 5. Review

Review communication, narrative, density, visual continuity, factual sources,
layout, native editability, and delivery. For a style Skill, compare the result
with its principles and relationships, not pixel similarity to examples. For a
reference deck, separate visual observation from source-preservation evidence.
For image-led or overlapping pages, inspect the final bottom-to-top stack and
confirm that evidence, labels, lines, markers, and foreground copy remain
unobstructed. Do not use a reference screenshot as the finished editable page.
Run editorial page-fit after the actual composition is visible, then commit the
reviewed candidate before follow-up edits or publication.
