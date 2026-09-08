# Design

## Contract

- `design.theme.accentTransforms` is an optional object keyed by
  `accent1` through `accent6`.
- Each present role may declare `tint` and/or `shade` in the inclusive range
  `0..1`; values use the same order as the existing PPJ color model: tint
  toward white, then shade toward black.
- The native wire stores each transform in thousandths of a percent, with
  optional presence for each operation.
- Authored output writes the transforms below the corresponding direct
  `a:srgbClr` accent owner. The embedded PPJ snapshot restores the exact
  authored values on projection.

## Boundaries

The profile does not transform arbitrary theme roles, resolve inherited
theme/effect schemes, edit imported `theme1.xml`, or claim host rendering
equivalence for every color-management path. Empty transform objects and
out-of-range/non-finite values fail closed.
