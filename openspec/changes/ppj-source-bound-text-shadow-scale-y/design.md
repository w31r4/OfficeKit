## Context

The native shadow model already preserves signed `sy` values, and the paragraph default-text path already proves and splices vertical scale tokens. Direct rich-text runs parse only the strict transform profile used by the existing horizontal-scale field, so a valid `outerShdw/@sy` makes the run opaque. The change remains source-bound and independent from shape/image shadow support.

## Goals / Non-Goals

**Goals:**

- Add one direct-run PPJ field with the existing signed Int32 / 100000 precision.
- Reuse the run index proof and token-splice path while accepting only one safe scale transform at a time.
- Keep RGB/theme color, geometry, alignment, opacity, effect ordering, and package residuals source-preserved.

**Non-Goals:**

- Editing skew, rotation, shape/image/chart shadows, or arbitrary effect graphs.
- Editing or adding a missing `sy` attribute.
- Making a claim about PowerPoint host rendering.

## Decisions

- Use the text-run-specific schema definition so the new field does not widen the shared shape/image `shadow` contract.
- Let the direct text reader parse at most one scale transform (`sx` or `sy`), then make projection and edit safety select the matching transform profile. A run with both scales remains opaque.
- Map `textShadowScaleY` to the existing `sy` XML attribute and preserve the surrounding XML by changing the token in place. Rebuilding the effect list would lose source spelling and unmodeled residuals.

## Risks / Trade-offs

- **Risk:** A run with multiple shadow transforms remains uneditable even when `sy` is present. → **Mitigation:** retain fail-closed checks for `sx` + `sy`, `kx`, `ky`, and `rotWithShape` and cover these variants in the focused regression.
- **Risk:** Scale values near the signed Int32 limits can be mis-rounded when authored. → **Mitigation:** use the existing 1/100000 ratio conversion and validate the native token before source-bound edits.
