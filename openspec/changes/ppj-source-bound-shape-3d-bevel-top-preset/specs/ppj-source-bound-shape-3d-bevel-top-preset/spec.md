# Shape 3-D top-bevel preset

## ADDED Requirements

### Requirement: Project one direct top-bevel preset

The system MUST expose one `shape3dBevelTopPreset` native leaf only for an
ordinary shape with exactly one direct `a:sp3d` whose only child is a direct
`a:bevelT` and whose `prst` value is a canonical token from the bounded
DrawingML bevel-preset vocabulary.

#### Scenario: Project a recognized top-bevel preset

- **WHEN** an imported shape contains a direct
  `a:sp3d/a:bevelT/@prst="angle"` with only bounded root and bevel attributes
- **THEN** PPJ exposes `shape3dBevelTopPreset` with value `angle`
- **AND** the projection preserves the other bevel and 3-D attributes as
  source state.

### Requirement: Edit only the top-bevel preset token

The system MUST allow a source-bound edit to replace only the proven
`a:bevelT/@prst` attribute in the owning SlidePart.

#### Scenario: Reproject an edited preset

- **WHEN** the `shape3dBevelTopPreset` leaf changes from `angle` to `softRound`
- **THEN** only the owning SlidePart changes
- **AND** the edited package remains Open XML-valid
- **AND** a second projection reports `softRound`.

### Requirement: Keep complex bevel state source-owned

The system MUST withhold the leaf and reject the source-bound operation when
the `sp3d` owner is missing, malformed, duplicated, extension-bearing,
child-bearing beyond the single `bevelT`, or has an unknown bevel attribute,
noncanonical value, or out-of-vocabulary preset.

#### Scenario: Keep a complex bevel opaque

- **WHEN** an imported shape's `sp3d` has a bottom-bevel, color, extension, or
  unknown `bevelT` attribute in addition to the strict top-bevel owner
- **THEN** the projection emits no `shape3dBevelTopPreset` leaf
- **AND** a source-bound edit cannot target that preset as an independent scalar.
