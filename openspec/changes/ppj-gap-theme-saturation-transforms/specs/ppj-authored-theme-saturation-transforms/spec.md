# PPJ authored theme saturation transforms

## ADDED Requirements

### Requirement: expose accent saturation modulation and offset

The PPJ schema MUST accept optional `satMod` and `satOff` values for each
`accent1` through `accent6` transform. Authored compilation MUST lower a
present value to the matching direct DrawingML accent color owner without
changing the base accent role or unrelated transforms.

#### Scenario: authored saturation transforms survive native output

- **WHEN** `accent4` has a base color and declares `satMod: 0.7` and
  `satOff: -0.2`
- **THEN** the authored theme contains `a:accent4Color/a:srgbClr` with
  `a:satMod val="70000"` followed by `a:satOff val="-20000"`
- **AND** embedded PPJ projection recovers both transform values

### Requirement: preserve legacy transform fields

When `satMod` and `satOff` are omitted, existing accent transform output MUST
remain unchanged.

#### Scenario: other transforms remain intact

- **WHEN** `accent1` declares only `tint` and `shade`
- **THEN** compilation emits those transforms without adding saturation
  children

### Requirement: reject unsafe saturation transforms

The compiler MUST reject non-finite, out-of-range, or empty transform values.

#### Scenario: invalid saturation transform fails closed

- **WHEN** a role declares `satOff: 1.2`
- **THEN** PPJ validation or compilation fails before native theme output
