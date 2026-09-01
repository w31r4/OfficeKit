# Visual attention contract

Use this reference when creating a page from a brief, selecting a visual
direction, or materially redesigning an imported page. It prevents the agent
from spending its first decisions on shapes instead of communication. It does
not prescribe a palette, font, template, or layout.

## Before composing

Write one compact contract before adding elements. Record it in the existing
PPJ `intent`, `design.grammar`, and page-planning fields; do not invent a new
schema just for this checklist.

```text
Audience task: what must this audience know, believe, decide, or do?
One conclusion: what is the sentence the page must leave behind?
Primary carrier: what carries the conclusion—chart, image, diagram, text,
  table, native vector, or a deliberate mix?
Reading order: what should be seen first, second, and third?
Protected evidence: which bars, lines, labels, numbers, axes, sources,
  image subjects, or relationships may not be hidden or cropped?
Canvas occupancy: which regions are intentionally active, quiet, or reserved?
Layer order: what is background, scrim, evidence, annotation, and foreground?
Forbidden shortcut: which filler pattern would make this page generic?
```

The contract should fit in a few lines. If the conclusion or carrier is still
unclear, resolve that uncertainty before composing. If the page intentionally
uses sparse composition, say why the quiet area creates focus, separation, or
rhythm.

## Choose the carrier by relationship

- Use a chart when magnitude, change, distribution, or uncertainty is the
  claim. Give axes, units, labels, and annotations enough room to remain true.
- Use a diagram when sequence, hierarchy, causality, flow, or spatial
  relationship is itself the evidence. Every connector and node needs a role.
- Use an image or SVG when a real subject, place, product, identity, or
  atmosphere carries meaning. Keep the focal subject inside the crop and keep
  attribution or evidence boundaries visible.
- Use typography when the conclusion is the evidence. Establish hierarchy by
  scale, measure, line, and rhythm rather than by a wall of containers.
- Use a deliberate mix only when the objects explain one another. A chart,
  image, and text block that merely coexist is not an information relationship.

## Compose in attention order

1. Place the primary carrier and give it enough area to be read at final size.
2. Place the conclusion and the smallest supporting labels needed to interpret
   that carrier.
3. Add decision, comparison, or uncertainty annotations only where they
   change the reading.
4. Add source, assumption, and accessibility text without hiding evidence.
5. Add a background, scrim, mask, line, or small motif only when its layer
   role is explicit. Do not use decoration to repair an empty page.

Use `pages[].elements[]` as a true back-to-front stack. A foreground object
must never cover a bar, line, marker, error bar, number, axis, source, image
subject, or connector that carries the page's claim. For combo charts, reserve
separate label clearances instead of forcing a line through a data label or
making a bar transparent merely to hide a layering mistake.

## 1→10 edits

For an existing page, inspect before editing:

- current carrier and conclusion;
- current reading order and z-order;
- evidence that must remain stable;
- intentional quiet areas and background layers;
- whether the request is local or a global redesign.

For a local edit, keep every untouched element, source, crop, and layer in
place. Do not add a new accent, card, image, or global rewrite simply because
the edited object leaves more space. For a redesign, write a replacement
contract and review all affected pages for rhythm and design drift.

## Review questions

Answer these from the rendered page, not from PPJ alone:

- Can a reader state the page conclusion without guessing?
- Is the primary carrier visibly dominant and large enough?
- Does the reading order follow scale, position, and contrast?
- Is every protected evidence object visible and legible?
- Does each quiet region have a declared purpose?
- Is any object present only to fill space or imitate a template?
- Do background images, masks, scrims, and overlays preserve contrast?
- Did a local edit change anything outside its requested surface?

Warnings are appropriate for weak hierarchy, repeated rhythm, or low
occupancy. Treat occlusion, clipping, unreadable text, false evidence, and
changed non-target content as blocking issues.

## What this contract does not do

It does not force dark or light themes, a particular number of images, a
specific geometry vocabulary, or a fixed page template. User templates,
brand systems, imported design, and the deck's own grammar remain authoritative.
The contract only makes the reasoning that precedes composition explicit and
reviewable.
