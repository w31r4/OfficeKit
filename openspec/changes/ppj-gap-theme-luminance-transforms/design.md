# Design

## Contract

- `design.theme.accentTransforms` remains keyed by `accent1` through
  `accent6`.
- Each role may declare `lumMod` as a fraction in `0..1` and `lumOff` as a
  signed fraction in `-1..1`.
- The native wire stores each value in thousandths of a percentage, with
  independent optional presence.
- Authored output writes `a:lumMod` and `a:lumOff` below the matching direct
  `a:srgbClr` accent owner. Existing `tint` then `shade` ordering is retained;
  luminance transforms follow those fields.

## Boundaries

The profile does not transform arbitrary theme roles, read or edit imported
`theme1.xml`, infer effect schemes, or claim host color-management
equivalence. Empty transform objects and out-of-range/non-finite values fail
closed.
