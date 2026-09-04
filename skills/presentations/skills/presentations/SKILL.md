---
name: presentations
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
3. Read [the visual attention contract](references/visual-attention.md), then
   read only the focused references needed by the planned carriers: fonts,
   text, shapes, charts/tables, media/layers, components/templates, or motion.
   When a required image is not supplied by the user, brand, template, or
   source deck, also read [image sourcing](references/image-sourcing.md).
   For a cutout or background-replacement role, use the image `visualProfile`
   (`alphaPresent`, `subjectBounds`, `edgeQuality`, `shadowMode`) before choosing
   crop, mask, contrast, or native shadow treatment.
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

Edit typed fields only when the matching capability is issued. For exact
imported text, style, or geometry scalars, change only `value` on an existing
`nativeRef.leaves[]` entry; never invent or move its ID, kind, or expected
hash. Keep the copied source asset and its SHA-256 unchanged. Unsupported
mutations fail closed; do not rebuild, flatten, rasterize, or patch the package
to force success.

For a visual edit, preserve the page's existing communication contract before
changing elements: identify its current audience task, primary claim, carrier,
reading order, protected evidence, and layer order. A local edit changes only
the requested surface. Replace the contract only when the user asks for a
redesign or a global style change.

### Recover an OfficeKit-authored deck

Import the PPTX. When its embedded program and node map are valid, OfficeKit
restores the authored PPJ exactly. That PPJ remains authoritative even if an
external application changed native PPTX state. Build to a new output; never
overwrite the input.

### Continue durable work

Task state is optional. Add `--task <id>` only when immutable PPJ revisions,
review receipts, resume, or publication evidence are useful. Resume from the
latest verified PPJ revision. If its reported status is not `reviewed`, treat
it as working state and do not deliver its candidate. Do not attempt to restore
a JavaScript heap. Legacy `ctx.plan` Presentation tasks are reported as
unsupported rather than silently migrated.

Create and continue the task through data-only commands:

```bash
officekit tasks --new "Continue the imported presentation" --json
officekit ppj import input.pptx -o deck.ppj --task <task-id> --json
officekit ppj check deck.ppj --task <task-id> --json
officekit ppj build deck.ppj -o candidate.pptx --task <task-id> --json
officekit ppj review deck.ppj --task <task-id> --json

# In a fresh context:
officekit tasks <task-id> --json
officekit ppj resume <task-id> -o resumed/deck.ppj --json
```

The resumed PPJ is a new editable working copy with its bound source and assets.
The immutable task revision remains read-only. Re-inspect the resumed program
before another edit and record the new check, build, and review with the same
task ID.

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

Before creating `pages[].elements[]`, write a short visual attention contract
in the existing `intent`, `design.grammar`, and page-planning fields:

```text
audience task → one conclusion → primary carrier → reading order
→ protected evidence → deliberate canvas occupancy → layer order
```

Do not start with a shape helper or a decorative asset. The contract is a
decision aid, not a fixed visual style: colors, fonts, geometry, imagery, and
density remain free to follow the selected scenario, template, or design
system. Use the detailed checklist in [visual attention](references/visual-attention.md)
for new pages and material redesigns.

For charts, missing observations, or layered pages, follow the four executable
quality contracts in that reference. Do not deliver from a structural pass
alone: render at final size, repair visible collisions or occlusion locally,
and state `visualReview: unavailable` when inspection was not possible.
If a series has only two observations, use the sparse-observation route: an
independent endpoint comparison is the default; a shared-axis sparse line is
an explicit exception, never an automatic fallback.

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
officekit ppj resume <task-id> -o resumed/deck.ppj --json
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
