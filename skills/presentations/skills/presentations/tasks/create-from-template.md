# Create from a template or brand reference

Use this route when a PPTX template, brand system, or visual reference is the
authoritative design source.

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

Use mode `create-from-template` and source mode `template` or `design-system`.
Link the authoritative source through `artifactRef`. The deck-specific grammar
may narrow or name source roles; it must not invent unsupported template facts.

## 4. Generate new content

Use source-slide and source-component reuse where capabilities support it.
Export/reimport a pending source-derived slide before adding supported overlays
or making another bounded edit. Use native Layout placeholders when they match
the content job. Compose new editable objects when the source has no suitable
archetype.

Do not flatten opaque source content, rebuild the package, or substitute Grid
for a failed template operation.

## 5. Review fidelity and delivery

Review the generated deck against the staged source and active plan. Verify
source protection, template facts, package integrity, page design, and visual
continuity. Commit the result before any follow-up edit or publication.
