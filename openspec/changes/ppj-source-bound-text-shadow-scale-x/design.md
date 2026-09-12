## Context

The native shadow model already preserves signed `sx` values and the paragraph default-text path already proves and splices scale tokens. Direct rich-text runs still use the transform-free shadow reader/profile, so a valid `outerShdw/@sx` makes the run opaque. The change must remain source-bound and independent from shape/image shadow support.

## Goals / Non-Goals

**Goals:**

- Add one direct-run PPJ field with the existing signed Int32 / 100000 precision.
- Reuse the run index proof and token-splice path while accepting only one safe scale transform.
- Keep RGB/theme color, geometry, alignment, opacity, effect ordering, and package residuals source-preserved.

**Non-Goals:**

- Editing `scaleY`, skew, rotation, shape/image/chart shadows, or arbitrary effect graphs.
- Adding or removing a missing `sx` attribute.
- Making a claim about PowerPoint host rendering.

## Decisions

- Use a text-run-specific schema definition so the new field does not silently widen the shared shape/image `shadow` contract.
- Let the direct text reader parse transforms, then make projection/edit safety require `sx` as the only transform; alignment and existing scalar fields may coexist.
- Map `textShadowScaleX` to the existing `sx` XML attribute and preserve the surrounding XML by changing the token in place. Rebuilding the effect list would lose source spelling and unmodeled residuals.

## Risks / Trade-offs

- **Risk:** A run with multiple shadow transforms remains uneditable even when `sx` is present. → **Mitigation:** retain fail-closed checks for `sy`, `kx`, `ky`, and `rotWithShape` and cover these variants in the focused regression.
- **Risk:** Scale values near the signed Int32 limits can be mis-rounded when authored. → **Mitigation:** use the existing 1/100000 ratio conversion and validate the native token before source-bound edits.
