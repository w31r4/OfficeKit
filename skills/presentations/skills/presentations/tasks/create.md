# Create a new presentation

Use this route when the user wants a new deck and has not supplied an
authoritative template.

## 1. Resolve only material ambiguity

Infer audience, purpose, outcome, duration, language, and evidence from the
request and available files. Ask at most three questions only when an answer
changes the audience, conclusion, evidence, or design authority.

Lock facts before rewriting. Read
[audience-facing text editing](../references/audience-text-editing.md).

## 2. Select the authoring route

- Shipped route A: use the Grid scaffold while it remains the recorded default.
- Experimental compiler route C: select `authoring-compiler-v1`, read
  [design mechanisms](../references/design-mechanisms.md), and create a
  deck-specific grammar before drawing.
- Use Grid explicitly when the user or plan requests it.

Do not switch routes after a failure. Repair the selected route or report the
blocking evidence.

## 3. Write the durable plan

For route C, follow [the authoring-plan contract](../references/authoring-plan.md):

```js
const planDescriptor = await ctx.plan(plan);
```

Give every page one reader task, claim, evidence set, content budget, and
composition intent. Select zero to two mechanism packs and write the actual
palette, typography, spacing, density rhythm, motif, imagery, chart treatment,
and strict invariants into `designGrammar`.

## 4. Compile a complete working draft

Use Presentation Help by intent, then compose with the golden primitives:

```js
const { Presentation, PresentationFile, reviewArtifact } = await ctx.import("office-kit");
const deck = Presentation.create();
deck.help("compose a planned presentation page");
```

Create theme/Master/Layout assets only when they improve reuse. Use
`slide.compose`, `slide.autoLayout`, native objects, and free coordinates as
the page intent requires. Keep page helpers reusable inside the task.

Build the complete working draft before asking the user to choose internal
layouts. The first response should offer the checked deck and a short summary
of its story, not a design questionnaire.

## 5. Review and commit

Follow [Review and deliver](review-deliver.md). Pass `authoringPlan` to
`reviewArtifact`, commit the reviewed candidate, then let the user continue by
conversation or accept delivery. For a self-directed deck, the final review
must be independent of any earlier draft: do not pass a source `baseline` to
hide errors created by the same task. Resolve semantic/layout errors such as
text overflow and unexpected overlap before committing; a baseline is for
source-bound local edits whose unchanged findings are already evidenced.
