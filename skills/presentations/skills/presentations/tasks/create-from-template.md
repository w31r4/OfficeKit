# Create from a template or brand reference

Use this route when a PPTX template, brand system, or visual reference is the
authoritative design source.

Before inspecting the source, read [presentation doctrine](../references/presentation-doctrine.md),
[the shared visual floor](../style_guidelines.md), [scenario policy](../references/scenario-policy.md),
and the selected primary scenario guide. Record the communication job, expected
audience change, delivery mode, after-use, and medium fit. A weak medium fit is
documented and mitigated; it does not silently change the requested deliverable.

Load the sibling
[`presentation-editorial-trim`](../../presentation-editorial-trim/SKILL.md)
Skill. Use its pre-composition pass after locking source wording and its
post-render pass after the template composition is visible. Template wording,
terminology, and factual qualifiers remain authoritative unless the requested
scope changes them.

## 1. Stage and identify authority

Copy the source with `ctx.input`. Record its artifact ID and SHA-256 in
`artifactRefs`; never rely on an absolute template path inside the authoring
plan. Keep the source read-only.

If several references exist, assign each one a role. Do not combine unrelated
design languages without an explicit user request.

## 2. Distill evidence

Import the PPTX, then use:

```js
const profile = presentation.designProfile({ maxItems: 64 });
const generation = presentation.planTemplateGeneration({ /* content needs */ });
```

Record evidenced palette, typography, spacing, density, archetypes, reusable
components, image/SVG assets, and unresolved decisions. The profile describes
the source; capability inspection decides what may be reused or edited.

Read [template-conditioned generation](../references/template-conditioned-generation.md)
for source-derived clone and continuation boundaries.

## 3. Write the plan

Use mode `create-from-template` and source mode `template`, `design-system`, or
`style-transfer` according to the actual authority. A user-supplied template or
brand system overrides scenario defaults; the scenario guide fills only
undefined decisions.
Link the authoritative source through `artifactRef`. The deck-specific grammar
may narrow or name source roles; it must not invent unsupported template facts.

Record one selected direction that explains how this deck will use the source
for its audience and content. Complete the grammar with palette/surface roles,
type rhythm, geometry and line rules, density rhythm, visual carriers,
image/SVG/chart/diagram treatment, allowed motifs, and anti-patterns. Mark
unresolved source facts instead of guessing them.

Every page `compositionIntent` must name its dominant carrier and source
strategy, such as a cloned source slide, reusable template component, supplied
image, source SVG, or newly authored native chart/diagram/table. “Follow the
template” is not a sufficient composition intent.

Template images and reusable source assets remain the first choice. Load
[image sourcing](../references/image-sourcing.md) only when the plan names a
media role that the authoritative source cannot fill. Record the registered
asset and provenance without presenting the external image as a template fact.

## 4. Generate new content

Use source-slide and source-component reuse where capabilities support it.
Export/reimport a pending source-derived slide before adding supported overlays
or making another bounded edit. Use native Layout placeholders when they match
the content job. Compose new editable objects when the source has no suitable
archetype.

Before expanding a deck longer than four pages, render an opening page, one
evidence page, and the densest or highest-risk page. Compare that spread with
the source's evidenced grammar. For four pages or fewer, inspect the complete
deck. If the grammar or reuse strategy changes, update the same authoring plan
with `expectedSha256`; do not create a second design state.

Do not flatten opaque source content, rebuild the package, or substitute Grid
for a failed template operation.

## 5. Review fidelity and delivery

Review the generated deck against the staged source and active plan. Verify
source protection, template facts, package integrity, page design, and visual
continuity. Run the editorial page-fit pass without erasing template voice or
source qualifiers. Commit the result before any follow-up edit or publication.
