---
name: "Presentations — Concise scenario route (experiment)"
description: "Create or edit editable PowerPoint presentations through PPJ with a short route, one scenario guide, and a task-local visual style brief."
---

# Presentations — concise route

Use the shared control contract first: [invariants](../../common/invariants.md).
Do not omit its evidence, occlusion, sparse-data, PPJ, source-bound, rights, or
render-repair rules when keeping this route short.

## Choose the job

Decide whether the request is **create**, **template/reference**, **import/edit**,
or **continue**. Name the audience, the one-sentence claim, the expected
change, delivery mode, and evidence boundary. Choose exactly one primary
scenario from [the scenario contract](../../common/what-kind.md), then read only
that guide from `../../common/references/scenarios/`.

For an edit, classify the requested scope before touching geometry:

- **local**: preserve the existing page contract, visual grammar, and non-target
  objects; change only the issued typed/nativeRef fields and then repair the
  smallest responsible visual layer;
- **reframe**: the user asks for a new claim, reading order, or visual direction,
  or the existing page cannot express the requested change safely. Re-compose
  the target page deliberately, but keep the source binding, stable IDs, and
  every non-target page/part intact.

Never silently turn a local edit into a reframe. Do not rebuild an entire
imported deck to make one target page look nicer.

## Make a local style decision

Before writing elements, create the task-local artifact described in
[style brief](../../common/style-brief.md). Let the claim and relationship
choose the carrier. If a template, brand system, or reference deck is supplied,
observe it and preserve its authority; otherwise choose a self-directed visual
direction. Do not turn the brief into a page skeleton or a reusable card
library.

Commit to one visual thesis, one primary carrier, and one reading order. Use the
canvas intentionally: a full-bleed image, background/mask stack, oversized type,
or a generous field of native geometry is valid when it carries the claim;
empty space must create focus rather than hide missing work. Do not default to
equal panels or a generic container merely because the page has several facts.

## Compose or edit in PPJ

Read [PPJ](../../common/references/ppj.md), then load only the references needed
by the selected carrier: fonts/text, shapes/lines, charts/tables,
image/layers, motion, templates, or imported native references. Write one strict
`.ppj` program with stable IDs and true element order. Never use MJS/JSX, raw
OOXML, XPath, relationship IDs, or another authoring engine.

For images, use the shared `officekit image` route. The query, candidate choice,
crop, rights, and alt text are part of the page decision; never use imagery as
filler. For imported PPTX, use `officekit ppj import`, edit only the issued
typed/nativeRef capability, and preserve opaque and non-target content.

## Check and deliver

```text
check → build → render at final size → repair visible defects → review → deliver
```

No bar, line, marker, label, number, source, mask, or background may obscure
another evidence-bearing object. A render or review failure requires a local
repair, not extra decoration. Inspect the highest-density intersections first:
labels against marks, connectors against nodes, foreground text against images,
and masks against contrast-critical text. If a collision remains, move or
recompose the responsible layer; never cover it with a new shape. Return
PPJ/PPTX paths, hashes, review status, and honest playback evidence. This is an
experimental arm; it does not choose a default.
