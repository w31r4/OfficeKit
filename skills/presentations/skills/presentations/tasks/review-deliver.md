# Review and deliver a presentation

Use this route after every meaningful creation or edit batch and before final
publication.

## Review in order

1. Factual: verify claims, sources, locked text, and traceability.
2. Communication: verify the audience change, job, scenario, delivery mode,
   medium-fit mitigation, and after-use.
3. Narrative: verify cumulative page responsibilities, evidence order, and the
   closing decision or action.
4. Cognitive and visual: inspect content budgets, hierarchy, density rhythm,
   visual carriers, composition, and plan-bound design risks.
5. Native operation: reimport, inspect the package, run layout/render and
   motion checks, and preserve source-bound topology.
6. Delivery: complete visual or human review when available, use AnyDoc only
   for a declared text/table coverage gap, then verify destination, source
   protection, bytes, and SHA-256.

Use only a renderer that the current installation explicitly provides. A clean
packed install does not imply Playwright, Chromium, or another optional visual
runtime is present; do not install one as part of the task. When it is absent,
complete the structural review and report `visualReview: "unavailable"`.

Read [design review](../references/design-review.md) for invariant and warning
semantics.

## Bind the plan and local scope

```js
const plan = await ctx.plan();
const review = await reviewArtifact(candidate, {
  authoringPlan: plan,
  changedPageIds,
  baseline: reviewedPath,
  outputPath: candidatePath,
  visualReview: "unavailable",
});
```

Omit `changedPageIds` for a full creation or explicit deck-wide redesign. Use
the exact plan page IDs for a local edit.

## Resolve findings

- Fix semantic, package, layout, strict plan, content-budget, required
  unresolved, and undeclared-page errors before commit.
- Evaluate repetition, density, rhythm, card-wall, dominant-geometry, hollow-
  container, text/container hierarchy, title-form, and design-drift warnings in
  context. They are deterministic signals, not aesthetic verdicts.
- Report `visualReview: "unavailable"` when the Agent cannot understand rendered
  pages. Structural checks do not become visual approval.

## Commit and publish

Commit only a non-failing review whose delivery hash matches the candidate.
Changing the plan after commit blocks publication until another reviewed
commit binds the new plan.

Publish with `ctx.publish(ctx.task.commit, { name })`. Return the absolute file
path, PPTX type, SHA-256, useful slide locators, evidence paths, limitations,
and visual-review status.
