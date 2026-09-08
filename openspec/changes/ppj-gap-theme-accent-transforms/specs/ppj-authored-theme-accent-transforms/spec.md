# PPJ authored theme accent transforms

## ADDED Requirements

### Requirement: expose accent tint and shade

The PPJ schema MUST accept optional `tint` and `shade` values for each
`accent1` through `accent6` transform. Authored compilation MUST lower a
present value to the matching direct DrawingML accent color owner without
changing the base accent role or unrelated theme roles.

#### Scenario: authored accent transform survives native output

- **WHEN** `accent1` has a base color and declares `tint: 0.25` and
  `shade: 0.1`
- **THEN** the authored theme contains `a:accent1Color/a:srgbClr` with
  `a:tint val="25000"` followed by `a:shade val="10000"`
- **AND** embedded PPJ projection recovers both transform values

### Requirement: preserve legacy accents

When `accentTransforms` is omitted, authored theme output MUST preserve the
existing accent color behavior and default palette.

#### Scenario: old program remains unchanged

- **WHEN** a PPJ declares only `theme.accentColors`
- **THEN** compilation emits the base accent colors without transform
  children

### Requirement: reject unsafe transforms

The compiler MUST reject non-finite, out-of-range, or empty transform values.

#### Scenario: invalid transform fails closed

- **WHEN** a role declares `tint: 1.2` or an empty transform object
- **THEN** PPJ validation or compilation fails before native theme output
