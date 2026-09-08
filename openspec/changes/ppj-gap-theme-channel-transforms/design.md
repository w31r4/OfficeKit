# Design

## Contract

- `design.theme.accentTransforms` remains keyed by `accent1` through
  `accent6`.
- Each role may declare the six channel fields independently:
  `redMod`/`greenMod`/`blueMod` use `0..1`; `redOff`/`greenOff`/`blueOff` use
  signed `-1..1` fractions.
- The native wire stores each value in thousandths of a percentage, with
  independent optional presence.
- Authored output writes the corresponding direct DrawingML channel
  transform children below the matching `a:srgbClr` accent owner after the
  existing tint/shade/luminance/alpha/saturation fields.

## Boundaries

The profile does not transform arbitrary theme roles, read or edit imported
`theme1.xml`, infer effect schemes, or claim host color-management
equivalence. Empty transform objects and out-of-range/non-finite values fail
closed.
