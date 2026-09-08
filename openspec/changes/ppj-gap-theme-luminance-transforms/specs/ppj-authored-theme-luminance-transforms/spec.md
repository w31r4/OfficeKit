# PPJ authored theme luminance transforms

## ADDED Requirements

### Requirement: expose accent luminance modulation and offset

The PPJ schema MUST accept optional `lumMod` and `lumOff` values for each
`accent1` through `accent6` transform. Authored compilation MUST lower a
present value to the matching direct DrawingML accent color owner without
changing the base accent role or unrelated transforms.

#### Scenario: authored luminance transforms survive native output

- **WHEN** `accent2` has a base color and declares `lumMod: 0.8` and
  `lumOff: -0.05`
- **THEN** the authored theme contains `a:accent2Color/a:srgbClr` with
  `a:lumMod val="80000"` followed by `a:lumOff val="-5000"`
- **AND** embedded PPJ projection recovers both transform values

### Requirement: preserve existing accent transforms

When `lumMod` and `lumOff` are omitted, existing accent `tint`/`shade` output
MUST remain unchanged.

#### Scenario: legacy transform fields remain intact

- **WHEN** `accent1` declares only `tint` and `shade`
- **THEN** compilation emits those transforms in their existing order without
  adding luminance children

### Requirement: reject unsafe luminance transforms

The compiler MUST reject non-finite, out-of-range, or empty transform values.

#### Scenario: invalid luminance transform fails closed

- **WHEN** a role declares `lumOff: 1.2`
- **THEN** PPJ validation or compilation fails before native theme output
