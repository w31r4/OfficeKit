# Shape 3-D bottom-bevel width

## ADDED Requirements

### Requirement: Project one direct bottom-bevel width

The system MUST expose one `shape3dBevelBottomWidthEmu` native leaf only for
an ordinary shape with exactly one direct `a:sp3d` whose only child is a direct
`a:bevelB` and whose `w` value is a canonical non-negative 32-bit coordinate.

#### Scenario: Project a recognized bottom bevel

- **WHEN** an imported shape contains a direct `a:sp3d/a:bevelB/@w="1200"`
  with only bounded root and bevel attributes
- **THEN** PPJ exposes `shape3dBevelBottomWidthEmu` with value `1200`
- **AND** the projection preserves the other bevel and 3-D attributes as
  source state.

### Requirement: Edit only the bottom-bevel width token

The system MUST allow a source-bound edit to replace only the proven
`a:bevelB/@w` attribute in the owning SlidePart.

#### Scenario: Reproject an edited width

- **WHEN** the `shape3dBevelBottomWidthEmu` leaf changes from `1200` to `2400`
- **THEN** only the owning SlidePart changes
- **AND** the edited package remains Open XML-valid
- **AND** a second projection reports `2400`.

### Requirement: Keep complex bevel state source-owned

The system MUST withhold the leaf and reject the source-bound operation when
the `sp3d` owner is missing, malformed, duplicated, extension-bearing,
child-bearing beyond the single `bevelB`, or has an unknown bevel attribute,
noncanonical value, or out-of-range width.

#### Scenario: Keep a complex bevel opaque

- **WHEN** an imported shape's `sp3d` has a top-bevel, color, extension, or
  unknown `bevelB` attribute in addition to the strict bottom-bevel owner
- **THEN** the projection emits no `shape3dBevelBottomWidthEmu` leaf
- **AND** a source-bound edit cannot target that width as an independent scalar.
