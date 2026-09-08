# PPJ authored theme hue transforms

## ADDED Requirements

### Requirement: expose accent hue modulation and offset

The PPJ schema MUST accept optional `hueMod` and `hueOff` values for each
`accent1` through `accent6` transform. `hueMod` MUST be a fraction in
`0..1`; `hueOff` MUST be a finite degree value in `-360..360`. Authored
compilation MUST lower a present value to the matching direct DrawingML
accent color owner without changing the base accent role or unrelated
transforms.

#### Scenario: authored hue transforms survive native output

- **WHEN** `accent6` has a base color and declares `hueMod: 0.75` and
  `hueOff: 45`
- **THEN** the authored theme contains `a:accent6Color/a:srgbClr` with
  `a:hueMod val="75000"` and `a:hueOff val="2700000"`
- **AND** embedded PPJ projection recovers both transform values

### Requirement: preserve existing hue transform presence

When either hue field is omitted, the compiler MUST NOT synthesize that
DrawingML child. Other accent transform fields MUST remain independent.

#### Scenario: one hue field remains independent

- **WHEN** `accent1` declares only `hueMod: 0.5`
- **THEN** compilation emits `a:hueMod` but no `a:hueOff`

### Requirement: reject unsafe hue transforms

The compiler MUST reject non-finite or out-of-range hue values before native
theme output.

#### Scenario: invalid hue offset fails closed

- **WHEN** a role declares `hueOff: 361`
- **THEN** PPJ validation or compilation fails
