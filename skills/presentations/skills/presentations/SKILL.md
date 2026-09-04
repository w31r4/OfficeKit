---
name: Presentations
description: Create, edit, continue, review, and deliver local PowerPoint PPTX presentations. Use for new decks, template-conditioned decks, imported PPTX edits, and durable presentation tasks. Use powerpoint-live-control instead when the user explicitly targets the currently open desktop PowerPoint presentation.
---

# Presentations

Use OfficeKit to turn a request into an editable, reviewed PPTX. Run task
modules with `officekit run`; use `officekit repl` when work must survive a new
Agent context.

Never import or use `@oai/artifact-tool`. It is a different host runtime, not an
OfficeKit alias, and its output must not be attributed to OfficeKit.

Use the shared [workspace](../office-kit/references/workspace.md),
[REPL](../office-kit/references/repl.md), and
[review](../office-kit/references/review.md) contracts.

## Choose one route

Load only the route that matches the request:

| Request | Route |
|---|---|
| Create a new deck from a goal or outline | [Create](tasks/create.md) |
| Create with a selected style Skill, design system, or reference deck | [Create from template](tasks/create-from-template.md) |
| Modify an existing local PPTX | [Edit existing](tasks/edit-existing.md) |
| Continue a saved OfficeKit task | [Continue](tasks/continue.md) |
| Review, finalize, or deliver a candidate | [Review and deliver](tasks/review-deliver.md) |

Use `powerpoint-live-control` only for an open, possibly unsaved desktop deck.
Keep local-file work here. For Google Slides, create and verify a local PPTX,
then follow the [local handoff](routing/google_slides.md); cloud upload is a
separate host action.

## Keep these invariants

- Treat inputs, templates, references, and accepted revisions as read-only.
  Write every candidate and final file to a distinct path.
- Use the public `office-kit` package and one authoring engine throughout the
  task. Do not silently switch engines or rebuild an imported file elsewhere.
- Reopen every exported PPTX. Export success alone is not acceptance evidence.
- Preserve unsupported imported topology as opaque. Use an inspected,
  capability-issued edit or stop at that object.
- Keep facts, source data, citations, and locked user text unchanged unless the
  requested scope includes them.
- Limit a local edit to its declared pages. A deck-wide editorial or visual
  rewrite requires explicit scope.
- Keep objects natively editable when OfficeKit can represent them; never
  flatten a source object merely to make an edit succeed.
- Do not claim visual review when rendered pages were not understood.

## Follow one production spine

```text
define → plan → design → compose/edit → review → commit → deliver
```

`define` identifies the audience, communication job, expected change, facts,
constraints, and after-use. Ask only when a missing answer materially changes
the outcome, evidence, or design authority.

`plan` assigns every page one reader job, one primary claim, evidence, a content
budget, and a dominant visual carrier. Durable work records this in the existing
presentation authoring plan; its schema lives in
[authoring plan](references/authoring-plan.md).

`design` starts from authority. A design system outranks a conflicting Template
Skill. A reference deck supplies observation or source-bound continuation, not
catalog identity. A selected Template Skill supplies style guidance and visual
examples; use them to write a new deck-specific Design Grammar and compose
freely. With no authority or suitable template, use the self-directed C route.

For ordinary self-directed pages, card-based composition is forbidden. Do not
build card walls, equal panel grids, colored-side cards, decorative pills or
badges, or pages dominated by `box()`, `card()`, or `metricPanel()`. This rule
does not alter an authoritative card-based template, imported source design,
real product UI, table/chart/image frames, a page-scale organizing field, or an
explicit dashboard request. The exact boundary is in the
[shared visual floor](style_guidelines.md).

`compose/edit` uses only public OfficeKit capabilities. Search Help by intent
and load advanced references only for the object or workflow in use. API
examples prove callability, not visual quality; do not copy their palette,
helpers, or page silhouettes as a design source.

`review` follows [Review and deliver](tasks/review-deliver.md). Run semantic,
structural, layout/render, design, optional reading-view, visual/human, and
delivery checks in that order. Motion is considered only for `live`, `hybrid`,
or explicitly animated work and follows [Motion](references/motion.md).

`commit` binds the reviewed candidate, evidence, and active plan revision.
Resume from the latest reviewed artifact and re-inspect it; never assume a new
context can recover the old JavaScript heap. The authoritative continuation
steps are in [Continue](tasks/continue.md).

## Load detail only when needed

- Creation loads the doctrine, visual floor, scenario policy, and one selected
  scenario guide.
- Catalog-style creation loads only the selected Template Skill, its examples,
  and the creation workflow. Reference-deck continuation loads its source-bound
  workflow instead.
- Image-led pages and any cross-type overlap load
  [layered composition](references/layered-composition.md).
- Imported editing loads advanced guidance only for the targeted native object;
  a z-order request also loads layered composition. Dense third-party files may
  additionally load the [complex imported-deck route](references/six-sample-import.md).
- Motion loads only for a speaking or explicit animation requirement.
- Review and continuation each have one authoritative task document; do not
  copy their mechanics into another route.

Avoid preloading the complete API, all scenario guides, every template, and all
advanced object references. Progressive loading is part of the workflow.

## Deliver evidence, not process noise

Return:

- the final absolute file path;
- `kind: "presentation"`;
- SHA-256;
- useful slide or object locators when stable;
- render, inspect, or verify evidence paths when available;
- the exact visual-review state: `complete`, `unavailable`, or
  `requires-human`.

Do not present scratch plans, previews, temporary builders, or QA files as the
deliverable. Call a deck final only after the selected route and review pass.
