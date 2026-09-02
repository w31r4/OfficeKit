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

## Make a local style decision

Before writing elements, create the task-local artifact described in
[style brief](../../common/style-brief.md). Let the claim and relationship
choose the carrier. If a template, brand system, or reference deck is supplied,
observe it and preserve its authority; otherwise choose a self-directed visual
direction. Do not turn the brief into a page skeleton or a reusable card
library.

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
repair, not extra decoration. Return PPJ/PPTX paths, hashes, review status, and
honest playback evidence. This is an experimental arm; it does not choose the
production default.
