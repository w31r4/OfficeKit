# Design

## Contract

- `design.theme.accentTransforms` remains keyed by `accent1` through
  `accent6`.
- Each role may declare `satMod` in `0..1` and signed `satOff` in `-1..1`.
- The native wire stores each value in thousandths of a percentage, with
  independent optional presence.
- Authored output writes `a:satMod` and `a:satOff` below the matching direct
  `a:srgbClr` accent owner after the existing luminance fields.

## Boundaries

The profile does not transform arbitrary theme roles, read or edit imported
`theme1.xml`, infer effect schemes, or claim host color-management
equivalence. Empty transform objects and out-of-range/non-finite values fail
closed.
