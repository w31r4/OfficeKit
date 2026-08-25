---
name: Presentations
description: Create, edit, continue, review, and deliver local PowerPoint PPTX presentations. Use for new decks, template-conditioned decks, imported PPTX edits, and durable presentation tasks. Use powerpoint-live-control instead when the user explicitly targets the currently open desktop PowerPoint presentation.
---

# Presentations

Use OfficeKit to turn a request into an editable PPTX, preserve every source
artifact, and return a reviewed working file. Run JavaScript with
`officekit run` or use the durable task REPL for work that will continue across
sessions.

Never import or use `@oai/artifact-tool`: it is a different host-bundled runtime,
not an OfficeKit alias or fallback, and its output must never be attributed to OfficeKit.

Use the installed OfficeKit contracts at
`../office-kit/references/workspace.md`, `../office-kit/references/repl.md`, and
`../office-kit/references/review.md` for workspace, durable-session, and shared
post-edit evidence rules.

## Route the request

Choose exactly one task route before loading detailed instructions:

| Request | Route |
|---|---|
| Create a new deck from a goal or outline | [Create](tasks/create.md) |
| Create new content under a supplied template or brand reference | [Create from template](tasks/create-from-template.md) |
| Modify an existing local PPTX | [Edit existing](tasks/edit-existing.md) |
| Continue a saved OfficeKit task | [Continue](tasks/continue.md) |
| Review, finalize, or deliver a candidate | [Review and deliver](tasks/review-deliver.md) |

Use `powerpoint-live-control` for an open, possibly unsaved desktop deck. Keep
local-file work in this Skill. Google Slides requests receive a verified PPTX
for local upload; follow [the local handoff route](routing/google_slides.md).

## Protect the artifact

- Treat every input, template, and reference as read-only. Write candidates and
  final files to distinct paths.
- Use the public `office-kit` package and `officekit run`. Keep one authoring
  engine and one source of artifact identity throughout the task.
- Reopen every exported PPTX before delivery. A successful export alone is not
  acceptance evidence.
- Preserve unsupported imported topology as opaque. Use capability-issued edits
  or stop at the affected object.
- Keep factual claims, source data, citations, and locked user text unchanged
  unless the requested scope includes them.
- A local edit changes only the declared pages. A deck-wide editorial or visual
  rewrite requires explicit scope.

## Use a durable authoring plan

For multi-step creation, broad redesign, template work, or any task likely to
continue in a fresh Agent context, use `officekit repl` and persist one
`office-kit/presentation-authoring-plan/v1` plan with `ctx.plan()`.

The plan records:

- mode, primary communication job, expected audience change, medium fit, and
  after-use;
- narrative thesis, sections, and ordered page jobs;
- one primary scenario, at most one secondary scenario, a chosen visual
  direction, an authoritative design source, and a deck-specific grammar;
- zero to two design mechanism packs;
- page claims, evidence, content budgets, composition intents, and assets;
- editorial rules, unresolved decisions, recipe, and next action.

Read [the authoring-plan contract](references/authoring-plan.md) before writing
or updating it. A plan update uses the current plan SHA-256. Review and commit
the next candidate before publishing under the new plan revision.

## Start creation from the communication task

For `create` and `create-from-template`, read these in order before composing:

1. [presentation doctrine](references/presentation-doctrine.md);
2. [the shared visual floor](style_guidelines.md);
3. [scenario policy](references/scenario-policy.md);
4. exactly one primary scenario guide, plus one secondary guide only when the
   plan records a real secondary scenario;
5. the selected design-source route and any chosen mechanisms.

The seven scenario guides are `scenario-analysis-decision.md`,
`scenario-business-proposal.md`, `scenario-management-report.md`,
`scenario-academic-research.md`, `scenario-education-training.md`,
`scenario-technical-engineering.md`, and `scenario-brand-creative.md`, all in
`references/`.

Record `brief.primaryJob`, `brief.expectedOutcome`, `brief.mediumFit`,
`design.scenario`, and `design.direction` before drawing. A weak medium fit is
not a stop condition: record the limitation and mitigation, then continue.

## Choose a design source

Record exactly one source mode in the plan:

- `self-directed`: derive a fresh visual language from audience, content, and
  communication goal.
- `design-system`: apply an explicit brand system supplied by the user.
- `template`: distill an authoritative PPTX and create within its language.
- `style-transfer`: use a visual reference as evidence while creating new,
  editable Office content.

User templates and brand rules take precedence over general presentation
preferences. Inspect them before choosing colors, fonts, layouts, or assets.

For a self-directed deck, select zero to two mechanisms from
[design mechanisms](references/design-mechanisms.md), then write concrete
choices for this deck. Mechanisms guide composition and rhythm; the plan owns
the actual palette roles and surface hierarchy, typography roles and size
rhythm, geometry and corner rules, line semantics, density rhythm, visual
carriers, imagery/SVG/chart/diagram treatment, allowed motifs, invariants, and
anti-patterns.

When no authority supplies a direction, form two or three materially different
directions internally and select one. Record its name and rationale. Ask the
user only when design authorities conflict or the choice changes a material
requirement.

The authoring-compiler recipe uses per-deck grammar and free Compose and is the
default for an unspecified self-directed deck. Grid Layout remains an explicit
scaffold: use it when the user or plan requests the Grid library, or when a
known compatibility fallback is required. A supplied template or design
system always takes precedence over either self-directed route. A failed route
stays failed or is repaired in place; it does not silently switch design
systems.

## Plan content before drawing

Ask questions only when the answer materially changes the audience, decision,
evidence, or authoritative design source. Ask no more than three at once.

Before authoring:

1. Lock facts, sources, user-provided wording, and constraints.
2. Load the sibling
   [`presentation-editorial-trim`](../presentation-editorial-trim/SKILL.md)
   Skill and shape titles, visible support, labels/sources, and notes before
   composition.
3. State the central conclusion or reader outcome.
4. Give every page one reader task and one primary claim.
5. Attach evidence to the page that uses it.
6. Set a realistic text and object budget for each page.
7. Choose a composition intent before coordinates or layout helpers. Name both
   the dominant carrier and whether it comes from a user/template asset,
   sourced/generated image, native chart/diagram/table, typography, vector, or
   a deliberate mix.
8. Write the plan, then compile it with existing Presentation primitives.

After rendering, run the editorial Skill's page-fit pass before deck-wide voice
review. For imported decks, it preserves the existing voice and only changes
declared pages unless the user explicitly requests a global rewrite.

For a self-directed deck longer than four pages, calibrate the opening page, an
evidence page, and the densest or highest-risk page before expanding the rest.
Use the complete deck when it has four pages or fewer. Render the spread, repair
the grammar in the same plan revision chain, then continue; do not create a
second design state or silently switch routes.

## Compile with existing primitives

Use Help by intent instead of scanning the entire API:

```js
const { Presentation } = await ctx.import("office-kit");
const deck = Presentation.create();
deck.help("compose a page from a planned content hierarchy");
deck.help("reuse one component from an imported template");
```

The golden creation surface includes:

- theme, Master, Layout, and placeholder discovery/use;
- `slide.compose` and `slide.autoLayout`;
- native shapes, connectors, images, tables, and charts;
- `presentation.validateLayout`;
- `presentation.designProfile()` and
  `presentation.planTemplateGeneration()`;
- `presentation.inspect()` and `resolve()`;
- source slide/component resolve, reuse, and continued editing;
- bounded SVG text and style edits.

Use free coordinates when they express the plan more clearly than a registered
layout. Prefer a small set of reusable helpers inside the task over copying
whole page implementations. Keep all objects selectable and editable unless a
source asset must remain opaque for fidelity.

## Use motion as a communication primitive

Finish the static page before adding motion:

```text
Narrative -> theme tokens -> page archetype -> visual carrier
          -> composition -> motion units -> review
```

Record `brief.deliveryMode` as `live`, `reader`, or `hybrid`, and
`design.motionPolicy` as `adaptive`, `none`, or `explicit`. Every page
`compositionIntent` must name its actual visual carrier. Motion may reveal a
relationship already present in the composition; it must not disguise an
under-composed page.

Choose motion from the speaking job, not from an effect catalogue. Reader
decks stay static. Hybrid decks use a short common transition and animate only
data, sequence, or causality. Live decks may use deliberate beats, usually no
more than four motion units per page.

```js
slide.animations.add(chart, {
  effect: "wipe", direction: "up", chartBuild: "category-element",
  start: "onClick", durationMs: 650, staggerMs: 90,
});
slide.animations.add(riskShape, {
  effect: "pulse", phase: "emphasis", start: "afterPrevious",
});
detailSlide.setMorph({
  from: overviewSlide,
  durationMs: 700,
  pairs: [{ key: "hero", from: overviewHero, to: detailHero }],
});
```

The supported effects are `fade`, `wipe`, `fly`, `zoom`, and `pulse`.
`textBuild` accepts `whole` or `paragraph`; chart builds accept
`all-at-once`, `series`, `category`, `series-element`, or `category-element`.
Use `withPrevious`, `afterPrevious`, or `onClick` to express order. Review by
inspecting the timing records, rendering every changed page, and checking the
static layout as well as the playback intent. Imported timing or Morph
extensions that are not capability-issued remain preserved and are not
reconstructed.

Use [the six motion recipes](references/motion.md) for data rise, causal
reveal, comparison beats, focus emphasis, calm continuity, and Morph
continuity. Use at most one baseline transition and one section transition in
a deck. Do not auto-advance or add sound.

Treat `image_view` and `image_generate` as optional Agent capabilities. Use
user/template images first, then native PowerPoint shapes, charts, tables, and
typography; use generated imagery when available and relevant. If `image_view`
is unavailable, keep structural QA and report `visualReview: "unavailable"`.
If a core visual cannot be understood or sourced, ask for that asset or deliver
an explicitly marked no-image version.

When a specific advanced imported object is requested, load
[advanced imported editing](references/advanced-imported-editing.md) and use
the inspected capability for that object. Reinspect after every source-bound
export because leaf IDs and expected hashes are revision-bound.

## Review after every meaningful edit

Follow [design review](references/design-review.md) and the selected task route.
The review order is:

1. semantic review;
2. structural package review;
3. layout/render review;
4. authoring-plan design checks;
5. optional text reading for a declared content-coverage gap;
6. visual or human review when available;
7. delivery review.

Pass the active plan and local edit scope to review:

```js
const review = await reviewArtifact(candidate, {
  authoringPlan: await ctx.plan(),
  changedPageIds: ["page-04"],
  baseline: reviewedPath,
  outputPath: candidatePath,
  visualReview: "unavailable",
});
```

Machine-provable plan violations can block delivery. Repetition, rhythm,
density, hierarchy, and title-form signals are review warnings; treat them as
prompts for judgment rather than aesthetic verdicts. When image understanding
is unavailable, report `visualReview: "unavailable"` and retain the structural
and design evidence.

Request AnyDoc only when visual review is unavailable or a long multi-page
inspection leaves a declared text/table coverage gap. It supplies a compact
text reading view and does not replace package, layout, image, or visual review.

## Commit, continue, and deliver

Use this durable sequence:

```text
tasks → repl → plan → compose/edit → review → commit → resume → publish
```

`ctx.commit` binds candidate SHA-256, review, Edit Plan evidence, and the active
authoring plan. Changing the plan returns the task to `working`; publishing
stays blocked until a new reviewed commit binds that revision.

A fresh session opens the latest reviewed artifact, reads `ctx.plan()`,
reimports the PPTX, and reinspects it. It does not restore JavaScript heap
state.

Return the final PPTX absolute path, file type, SHA-256, slide locators when
useful, review evidence paths, and the exact visual-review status. Call a file
final only after its selected review route passes.
