# Review and deliver a presentation

Use this route after every meaningful creation or edit batch and before final
publication.

## Review in order

1. Reimport the candidate and verify semantic content.
2. Inspect the OOXML package and relationships.
3. Run layout/render checks for overflow, bounds, overlap, crop, and geometry.
4. Run authoring-plan design checks.
5. Request an AnyDoc reading view only for a declared text/table coverage gap.
6. Complete visual or human review when available.
7. Verify destination, source protection, bytes, and SHA-256.

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
- Evaluate repetition, density, rhythm, card-wall, title-form, and design-drift
  warnings in context. They are deterministic signals, not aesthetic verdicts.
- Report `visualReview: "unavailable"` when the Agent cannot understand rendered
  pages. Structural checks do not become visual approval.

## Commit and publish

Commit only a non-failing review whose delivery hash matches the candidate.
Changing the plan after commit blocks publication until another reviewed
commit binds the new plan.

Publish with `ctx.publish(ctx.task.commit, { name })`. Return the absolute file
path, PPTX type, SHA-256, useful slide locators, evidence paths, limitations,
and visual-review status.
