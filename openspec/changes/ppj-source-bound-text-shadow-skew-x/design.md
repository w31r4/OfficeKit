## Context

`PresentationShadow` and `PptxShadowCodec.TryReadOuterShadow` already retain signed horizontal skew in 60000ths of a degree when transform parsing is enabled. Direct rich-text projection currently permits only one scale transform and deliberately rejects both skew attributes, so a source run containing only `outerShdw/@kx` loses its editable semantic owner. Paragraph default-text shadows already expose the analogous field; this change keeps the direct-run scope narrow and source-bound.

## Goals / Non-Goals

**Goals:**

- Add one direct-run PPJ field with the existing signed 60000ths-of-a-degree precision.
- Reuse the run index proof and token-splice path while accepting only one safe shadow transform at a time.
- Keep RGB/theme color, geometry, alignment, opacity, effect ordering, and package residuals source-preserved.

**Non-Goals:**

- Editing `skewY`, scale, rotation, shape/image/chart shadows, or arbitrary effect graphs.
- Editing or adding a missing `kx` attribute.
- Making a claim about PowerPoint host rendering.

## Decisions

- Use the text-run-specific schema definition so this field does not widen the shared shape/image `shadow` contract.
- Parse direct runs with transform attributes enabled, then require at most one of `sx`, `sy`, `kx`, or `ky`. Projection and edit safety select the matching transform profile; a run with multiple transforms remains opaque.
- Map `textShadowSkewX` to the existing `kx` XML attribute and preserve surrounding XML by changing the token in place. Rebuilding the effect list would lose source spelling and unmodeled residuals.
- Represent the semantic value as degrees and round authored values to the native 60000ths-of-a-degree integer with the existing angle helper. Native values remain strictly inside -90° and 90°.

## Risks / Trade-offs

- **Risk:** A run with several shadow transforms remains uneditable even when `kx` is present. → **Mitigation:** retain fail-closed checks for `sx`, `sy`, `ky`, `rotWithShape`, and sibling/unknown content, and cover these variants in the focused regression.
- **Risk:** Values near the strict skew limits can be mis-rounded. → **Mitigation:** validate the rounded native token after conversion and reject boundary values that would become ±5400000.
