# Design

## Contract

- `design.theme.accentTransforms` remains keyed by `accent1` through
  `accent6`.
- Each role may declare `gray`, `comp`, `inv`, `gamma`, and `invGamma` as
  explicit `true` values. These fields are independent and optional.
- The native wire preserves each field's presence as an optional boolean.
- Authored output writes `a:gray`, `a:comp`, `a:inv`, `a:gamma`, and
  `a:invGamma` below the matching direct `a:srgbClr` accent owner.

## Boundaries

The profile does not transform arbitrary theme roles, read or edit imported
`theme1.xml`, infer effect schemes, or claim host color-management
equivalence. False-valued and empty transform declarations fail closed at
schema validation; omitted operations are not synthesized.
