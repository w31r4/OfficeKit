## Context

See proposal.md. XlsxChartTextEffectsCodec owns chart glow/outer-shadow composition. PptxSoftEdgeCodec has direct radius validation plus separate Presentation graph rules referencing reflection; moving that whole codec into Shared would unnecessarily move graph-specific dependencies.

## Goals / Non-Goals

**Goals:** add one chart field while sharing direct radius semantics and preserving each existing effect.

**Non-Goals:** broadening ordinary imported effect graphs, radius tokens beyond the ordinary field contract, arbitrary native effects, renderer or PowerPoint host acceptance.

## Decisions

- Extract direct soft-edge validation/reading into a Shared PptxSoftEdgeValueCodec. Keep PptxSoftEdgeCodec's existing composition rules in Presentation and delegate its direct-value checks. This shares the value contract without moving reflection or changing ordinary topology policy.
- Add PresentationSoftEdge field 21 to chart text style and reuse the existing softEdge JSON definition. Include field presence in meaningful-style/global-font guards and semantic comparison.
- Extend the chart effects adapter to peel one final softEdge from a cloned list, prove its exact direct value, then validate the remaining glow/shadow prefix with the existing readers. Preserve strict text/comment/unknown-descendant checks. Write the same native order with one list.
- The authored three-effect experiment exposed an existing resource-reference mismatch: direct glow/shadow builders accept standard theme tokens, but generic resource validation rejects them before compilation. Permit recognized undeclared themes only at those effect color paths; retain wrong-kind grammar-token and RGB foreground rejection.
- Use a chart radius builder with finite/range checks and the existing EMU conversion in authored/source-bound/vector paths. Project via the existing soft-edge representation. Honor nested style rules and explicit vector run overrides.
- Reuse line/combo lifecycle and style-owner helpers, inspect native order and ZIP preservation, and retain prior glow/shadow tests plus ordinary soft-edge regressions.

## Risks / Trade-offs

- A previous negative glow-plus-softEdge fixture becomes valid → replace it with malformed radius and add positive coexistence coverage.
- Deletion might erase siblings or collapse zero to absence → exercise all three independent fields and fresh projection after each step.
- A shared helper extraction can change ordinary behavior → retain the old graph rules and run ordinary effect regressions.
