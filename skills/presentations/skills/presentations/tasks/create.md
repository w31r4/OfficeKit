# Create a new presentation

Use this route for a self-directed new deck. If template search selected a style
Skill, or the user supplies a design system or reference deck, switch to
[Create from template](create-from-template.md).

## 1. Define the communication task

Infer the audience, communication job, expected audience change, language,
duration, after-use, constraints, and evidence from the request and available
files. Ask at most three questions, and only when an answer changes the
conclusion, evidence, or design authority.

Lock facts, sources, quotes, and user wording before rewriting. Load the sibling
[`presentation-editorial-trim`](../../presentation-editorial-trim/SKILL.md)
Skill. Use its pre-composition pass now and its page-fit pass after rendering.

Read, in this order:

1. [Presentation doctrine](../references/presentation-doctrine.md);
2. [shared visual floor](../style_guidelines.md);
3. [scenario policy](../references/scenario-policy.md);
4. exactly one primary scenario guide, plus one secondary guide only when a
   named section truly serves a different audience task.

Record `primaryJob`, `expectedOutcome`, `mediumFit`, `afterUse`, and the selected
scenario. If slides are a weak fit, record the limitation and mitigation, then
continue rather than pretending the medium has no boundary.

## 2. Choose the design source

Choose one source:

- `self-directed`: invent a visual language for this deck;
- `design-system`: apply supplied brand rules;
- `template`: use one schema-v3 Template Skill and its visual examples;
- `style-transfer`: use a reference deck as visual evidence while producing
  new editable Office content.

User authority wins. This route normally uses `self-directed`; the other three
sources move to the template route when their authority is material. A
reference deck that must remain the actual source uses source continuation,
not `template`.

For self-directed work, read
[design mechanisms](../references/design-mechanisms.md), form two or three
materially different directions internally, choose one, and record why it fits
the audience, evidence, delivery mode, and after-use. Use zero to two mechanism
packs. They guide composition but do not supply a palette or template.

The C authoring route is the default. Write one deck-specific Design Grammar
covering palette roles, type roles, geometry and line behavior, density rhythm,
visual carriers, asset treatment, motifs, invariants, and forbidden patterns.

Ordinary self-directed pages must obey the hard card ban in the shared visual
floor. Resolve parallel information with scale, position, baseline, light
rules, numbering, connectors, charts, tables, imagery, and intentional negative
space instead of container grids.

## 3. Plan the deck

For durable work, write the existing
[authoring plan](../references/authoring-plan.md). Give every page:

- one reader task and one primary claim;
- evidence and source constraints;
- a realistic content budget;
- one dominant visual carrier and its asset/source strategy;
- a page silhouette that contributes to the deck-wide rhythm.

When the carrier is an image or the composition depends on overlapping object
types, read [layered composition](../references/layered-composition.md). Plan the
bottom-to-top relationship explicitly. An image-led page normally starts with
the image, adds only the contrast treatment it needs, and keeps editable copy,
evidence, and decisions in the foreground.

Use a quantitative relationship as a chart, axis, direct label, spatial
comparison, or diagram rather than an inventory of numbers. Title pages should
be minimal; executive summaries should scan quickly through hierarchy and
light dividers; closing pages resolve the opening and add no new claim.

Set a readability floor appropriate to the delivery. For the ordinary
self-directed route, start with `minimumBodyFontSize: 22` and
`minimumCaptionFontSize: 20`; shorten, split, or recompose before shrinking.
When there are six or more pages, plan at least four meaningful silhouettes.

Set `deliveryMode` to `live`, `reader`, or `hybrid`. Read
[Motion](../references/motion.md) only for `live`, `hybrid`, or an explicit
animation request, and only after the static composition is complete.

## 4. Calibrate before expanding

For a deck longer than four pages, first build:

1. the opening or visual-direction page;
2. one evidence or data page;
3. the densest or highest-risk page.

For four pages or fewer, build the complete deck. Render the calibration spread
and inspect it as one sequence. Confirm that the grammar works across sparse and
dense pages, each planned carrier actually owns its page, negative space has a
clear purpose, and text and visuals belong to one system.

Repair the same direction and update the same plan revision. Do not switch to a
different template or a second design state after a failed calibration.

## 5. Compose and review

Search Presentation Help by intent and use public OfficeKit primitives. Examples
show API and file workflows, not a house style; do not copy their palettes,
generic container helpers, or page silhouettes.

Use `slide.elements` and the shared ordering methods when composition depends on
layering. Do not rely on a type collection or source-code statement order as a
substitute for the exported scene stack.

Compose in this order:

```text
communication job → scenario → direction → design grammar
                  → page archetype → visual carrier → composition → motion
```

The static page must work before motion. Enlarge, crop, align, or recompose the
evidence carrier when the canvas feels hollow; do not fill accidental emptiness
with panels. Run `presentation.validateLayout()` before export, reopen the PPTX,
and then follow [Review and deliver](review-deliver.md).

Run the editorial page-fit and deck-voice passes without changing locked facts
or the selected visual direction. Resolve overflow, unintended overlap,
off-canvas content, weak contrast, and unrecorded design warnings before commit.
Offer the reviewed working deck and a concise story summary; do not ask the user
to select internal layouts before a complete draft exists.
