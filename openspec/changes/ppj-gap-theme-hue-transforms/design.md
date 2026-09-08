# Design

## Contract

- `design.theme.accentTransforms` remains keyed by `accent1` through
  `accent6`.
- Each role may declare `hueMod` in `0..1` and signed `hueOff` in degrees
  `-360..360`.
- The native wire stores `hueMod` in thousandths of a percentage and
  `hueOff` in signed 1/60000-degree units, with independent optional
  presence.
- Authored output writes `a:hueMod` and `a:hueOff` below the matching direct
  `a:srgbClr` accent owner.

## Boundaries

The profile does not transform arbitrary theme roles, read or edit imported
`theme1.xml`, infer effect schemes, or claim host color-management
equivalence. Non-finite or out-of-range values fail closed; omitted fields are
not synthesized.
