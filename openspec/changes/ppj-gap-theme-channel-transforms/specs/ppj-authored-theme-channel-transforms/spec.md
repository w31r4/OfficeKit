# PPJ authored theme channel transforms

## ADDED Requirements

### Requirement: expose accent RGB channel modulation and offset

The PPJ schema MUST accept optional `redMod`/`redOff`, `greenMod`/`greenOff`,
and `blueMod`/`blueOff` values for each `accent1` through `accent6`
transform. Authored compilation MUST lower a present value to the matching
direct DrawingML accent color owner without changing the base accent role or
unrelated transforms.

#### Scenario: authored RGB channel transforms survive native output

- **WHEN** `accent4` has a base color and declares `redMod: 0.8`,
  `redOff: -0.05`, `greenMod: 0.6`, `greenOff: 0.1`, `blueMod: 0.7`, and
  `blueOff: -0.2`
- **THEN** the authored theme contains `a:accent4Color/a:srgbClr` with the
  corresponding `a:redMod`, `a:redOff`, `a:greenMod`, `a:greenOff`,
  `a:blueMod`, and `a:blueOff` values in declaration order
- **AND** embedded PPJ projection recovers all six transform values

### Requirement: preserve existing accent transforms

When RGB channel fields are omitted, existing accent tint, shade, luminance,
alpha, and saturation output MUST remain unchanged; when only some channels
are present, omitted channel transforms MUST not be synthesized.

#### Scenario: partial channel transforms remain partial

- **WHEN** `accent1` declares only `redMod` and `blueOff`
- **THEN** compilation emits only those channel children in addition to any
  explicitly declared earlier transforms
- **AND** no green channel child is added

### Requirement: reject unsafe RGB channel transforms

The compiler MUST reject non-finite, out-of-range, or empty transform values.
Modulation values MUST be within `0..1`, and offset values MUST be within
`-1..1`.

#### Scenario: invalid channel transform fails closed

- **WHEN** a role declares `greenOff: 1.2`
- **THEN** PPJ validation or compilation fails before native theme output
