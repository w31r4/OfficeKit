# Shape 3-D top-bevel height

## ADDED Requirements

### Requirement: Project one direct top-bevel height

The system MUST expose one `shape3dBevelTopHeightEmu` native leaf only for an
ordinary shape with exactly one direct `a:sp3d` whose only child is a direct
`a:bevelT` and whose `h` value is a canonical non-negative 32-bit coordinate.

#### Scenario: Project a recognized top bevel height

- **WHEN** an imported shape contains a direct `a:sp3d/a:bevelT/@h="800"`
  with only bounded root and bevel attributes
- **THEN** PPJ exposes `shape3dBevelTopHeightEmu` with value `800`
- **AND** the projection preserves the other bevel and 3-D attributes as
  source state.

### Requirement: Edit only the top-bevel height token

The system MUST allow a source-bound edit to replace only the proven
`a:bevelT/@h` attribute in the owning SlidePart.

#### Scenario: Reproject an edited height

- **WHEN** the `shape3dBevelTopHeightEmu` leaf changes from `800` to `1600`
- **THEN** only the owning SlidePart changes
- **AND** the edited package remains Open XML-valid
- **AND** a second projection reports `1600`.

### Requirement: Keep complex bevel state source-owned

The system MUST withhold the leaf and reject the source-bound operation when
the `sp3d` owner is missing, malformed, duplicated, extension-bearing,
child-bearing beyond the single `bevelT`, or has an unknown bevel attribute,
noncanonical value, or out-of-range height.

#### Scenario: Keep a complex bevel opaque

- **WHEN** an imported shape's `sp3d` has a bottom-bevel, color, extension, or
  unknown `bevelT` attribute in addition to the strict top-bevel owner
- **THEN** the projection emits no `shape3dBevelTopHeightEmu` leaf
- **AND** a source-bound edit cannot target that height as an independent scalar.
