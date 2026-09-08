# Shape 3-D bottom-bevel preset

## ADDED Requirements

### Requirement: Project one direct bottom-bevel preset

The system MUST expose one `shape3dBevelBottomPreset` native leaf only for an
ordinary shape with exactly one direct `a:sp3d` whose only child is a direct
`a:bevelB` and whose `prst` value is in the bounded DrawingML bevel-preset
vocabulary.

#### Scenario: Project a recognized bottom bevel preset

- **WHEN** an imported shape contains a direct
  `a:sp3d/a:bevelB/@prst="angle"` with only bounded root and bevel attributes
- **THEN** PPJ exposes `shape3dBevelBottomPreset` with value `angle`
- **AND** the projection preserves the other bevel and 3-D attributes as
  source state.

### Requirement: Edit only the bottom-bevel preset token

The system MUST allow a source-bound edit to replace only the proven
`a:bevelB/@prst` attribute in the owning SlidePart.

#### Scenario: Reproject an edited preset

- **WHEN** the `shape3dBevelBottomPreset` leaf changes from `angle` to
  `softRound`
- **THEN** only the owning SlidePart changes
- **AND** the edited package remains Open XML-valid
- **AND** a second projection reports `softRound`.

### Requirement: Keep complex bevel state source-owned

The system MUST withhold the leaf and reject the source-bound operation when
the `sp3d` owner is missing, malformed, duplicated, extension-bearing,
child-bearing beyond the single `bevelB`, or has an unknown bevel attribute,
missing/unknown/noncanonical preset, or otherwise ambiguous state.

#### Scenario: Keep a complex bevel opaque

- **WHEN** an imported shape's `sp3d` has a top-bevel, color, extension, or
  unknown `bevelB` attribute in addition to the strict bottom-bevel owner
- **THEN** the projection emits no `shape3dBevelBottomPreset` leaf
- **AND** a source-bound edit cannot target that preset as an independent
  scalar.
