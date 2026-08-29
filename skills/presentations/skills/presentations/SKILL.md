---
name: Presentations
description: Create, import, edit, continue, render, review, and deliver editable PowerPoint presentations through OfficeKit PPJ. Use powerpoint-live-control instead only when the user explicitly targets the presentation currently open in desktop PowerPoint.
---

# Presentations

PPJ is OfficeKit's only public Presentation authoring language. A `.ppj` file
is one strict, non-executable JSON program. Edit PPJ directly, then compile it
with the native OfficeKit compiler:

```text
request or PPTX → deck.ppj → check → build → render → review → deliver
```

Read [the PPJ language reference](references/ppj.md) before writing or changing
a program. Do not use MJS, JSX, legacy composition/facade APIs, raw OOXML,
XPath, relationship IDs, or another authoring engine as a substitute.
An external script may generate JSON, but executable code is not part of PPJ.

Use `powerpoint-live-control` only for an open, possibly unsaved desktop deck.
For local PPTX and PPJ files, stay in this Skill. Never import
`@oai/artifact-tool` or attribute its output to OfficeKit.

## Choose one route

### Create

1. Define the audience, communication job, expected change, evidence boundary,
   delivery mode, and after-use.
2. Read [scenario routing](references/scenarios/README.md) and exactly one
   primary scenario guide.
3. Read only the focused references needed by the planned carriers: fonts,
   text, shapes, charts/tables, media/layers, components/templates, or motion.
4. Write one deck-specific Design Grammar and ordered page plan into PPJ.
5. Build, render, review, revise the PPJ, and emit a new PPTX path.

### Create from a template, design system, or reference

Read [components and templates](references/components-and-templates.md).
A design system is binding brand authority. A selected Template Skill supplies
style guidance and representative images. A reference PPTX may additionally
supply observed design evidence or a reusable source package. Do not mix
unrelated templates. Derive a deck-specific grammar and compose every page in
PPJ; a template guides design rather than pinning a page skeleton.

### Import or edit an existing PPTX

Read [imported native references](references/imported-native-ref.md), then run:

```bash
officekit ppj import input.pptx -o deck.ppj
officekit ppj inspect deck.ppj --query "target"
```

Edit typed fields or fields explicitly issued by `nativeRef`. Keep the copied
source asset and its SHA-256 unchanged. Unsupported mutations fail closed; do
not rebuild, flatten, rasterize, or patch the package to force success.

### Recover an OfficeKit-authored deck

Import the PPTX. When its embedded program and node map are valid, OfficeKit
restores the authored PPJ exactly. That PPJ remains authoritative even if an
external application changed native PPTX state. Build to a new output; never
overwrite the input.

### Continue durable work

Task state is optional. Add `--task <id>` only when immutable PPJ revisions,
review receipts, resume, or publication evidence are useful. Resume from the
latest verified PPJ revision and reopen its reviewed artifact. Do not attempt
to restore a JavaScript heap. Legacy `ctx.plan` Presentation tasks are reported
as unsupported rather than silently migrated.

### Review and deliver

Read [review and delivery](references/review-and-delivery.md). Compilation is
not rendering, rendering is not visual understanding, and structural playback
evidence is not desktop PowerPoint acceptance.

## Design before drawing

Use this order for every new deck:

```text
communication job
→ narrative and page responsibilities
→ design authority and deck-specific grammar
→ evidence relationship and visual carrier
→ composition and true layer order
→ optional motion
→ review
```

Each page needs one audience task, one primary claim, evidence, a content
budget, and a dominant relationship or explicit `none`. Choose the carrier
that communicates that relationship: text, image, chart, table, diagram,
native vector, or a deliberate mix. Negative space must create focus,
separation, or rhythm; it is not leftover canvas.

Arrays are semantic. `pages[]` is page order. `pages[].elements[]` is the real
back-to-front z-order. Use stable IDs and edit those IDs in later turns.

## Strictly forbidden

- Fabricating data, sources, citations, cases, experiments, or certainty.
  Mark missing material as a placeholder or explicit assumption.
- Building hierarchy from card walls, equal rounded panels, colored side-strip
  cards, pills, badges, or a universal `box()` pattern.
- Adding random circles, rings, arrows, icon clouds, gradients, or stock images
  merely to fill space.
- Putting red, purple, yellow, and green accents together as generic "AI"
  styling instead of defining controlled palette roles.
- Letting bars, fills, masks, labels, shapes, or decoration hide a line,
  marker, error bar, number, source, or other evidence-bearing object.
- Using tiny text to rescue an overloaded page, or sparse composition to hide
  missing evidence.
- Putting file-system escapes, remote URLs, code, functions, recursion,
  unbounded loops, or arbitrary expressions inside PPJ.

Authoritative user templates and untouched imported design remain
authoritative, but newly added objects still need a communication role and may
not obscure evidence.

## Common commands

```bash
officekit ppj inspect deck.ppj --json
officekit ppj check deck.ppj --json
officekit ppj build deck.ppj -o deck.pptx --json
officekit ppj render deck.ppj -o previews/ --json
officekit ppj review deck.ppj --json
```

`check --fix` may apply deterministic formatting fixes only. It must not invent
copy, select a visual direction, or alter design semantics.

## Deliver evidence

Return the absolute PPJ and PPTX paths, PPTX SHA-256, review status, and useful
render or inspection evidence. State visual review as `complete`,
`unavailable`, or `requires-human`. Keep source files, PPJ revisions, previews,
and scratch outputs distinct from the final deliverable.
