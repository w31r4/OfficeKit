# Design

## Contract

- `design.theme.accentTransforms` remains keyed by `accent1` through
  `accent6`.
- Each role may declare `alphaMod` in `0..1` and signed `alphaOff` in
  `-1..1`.
- The native wire stores each value in thousandths of a percentage, with
  independent optional presence.
- Authored output writes `a:alpha` for an RGBA base color first, then
  `a:alphaMod` and `a:alphaOff` below the matching direct `a:srgbClr` owner.

## Boundaries

The profile does not transform arbitrary theme roles, read or edit imported
`theme1.xml`, infer effect schemes, or claim host alpha-compositing
equivalence. Empty transform objects and out-of-range/non-finite values fail
closed.
