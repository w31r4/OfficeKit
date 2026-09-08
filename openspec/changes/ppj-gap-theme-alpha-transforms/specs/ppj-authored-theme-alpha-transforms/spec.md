# PPJ authored theme alpha transforms

## ADDED Requirements

### Requirement: expose accent alpha modulation and offset

The PPJ schema MUST accept optional `alphaMod` and `alphaOff` values for each
`accent1` through `accent6` transform. Authored compilation MUST lower a
present value to the matching direct DrawingML accent color owner without
changing the base accent role or unrelated transforms.

#### Scenario: authored alpha transforms survive native output

- **WHEN** `accent3` has a base color and declares `alphaMod: 0.6` and
  `alphaOff: -0.1`
- **THEN** the authored theme contains `a:accent3Color/a:srgbClr` with
  `a:alphaMod val="60000"` followed by `a:alphaOff val="-10000"`
- **AND** embedded PPJ projection recovers both transform values

### Requirement: preserve the absolute alpha owner

When the base accent is an RGBA color, adding `alphaMod` or `alphaOff` MUST
retain the direct `a:alpha` child before the relative alpha transforms.

#### Scenario: RGBA base remains distinct

- **WHEN** an accent has an eight-digit base color and an alpha transform
- **THEN** authored output retains both the absolute `a:alpha` and the declared
  relative alpha transform children

### Requirement: reject unsafe alpha transforms

The compiler MUST reject non-finite, out-of-range, or empty transform values.

#### Scenario: invalid alpha transform fails closed

- **WHEN** a role declares `alphaOff: 1.2`
- **THEN** PPJ validation or compilation fails before native theme output
