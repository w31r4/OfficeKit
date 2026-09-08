# Picture shape 3-D depth

## ADDED Requirements

### Requirement: Project one direct picture depth scalar

The system MUST expose one `shape3dDepthEmu` native leaf for an imported
picture only when its `p:spPr` contains one child-free direct `a:sp3d` with a
canonical signed 32-bit `z` coordinate.

#### Scenario: Project a negative picture depth

- **WHEN** an imported picture contains direct `p:spPr/a:sp3d/@z="-1200"`
- **THEN** PPJ exposes `shape3dDepthEmu` with value `-1200`
- **AND** the projection preserves the picture asset, mask, effects, and
  other direct 3-D attributes as source state.

### Requirement: Edit only the picture depth token

The system MUST allow a source-bound `shape3dDepthEmu` edit to replace only
the proven picture `sp3d/@z` token in the owning SlidePart.

#### Scenario: Reproject an edited picture depth

- **WHEN** the picture leaf changes from `-1200` to `-2400`
- **THEN** only the owning SlidePart changes
- **AND** the edited package remains Open XML-valid
- **AND** a second projection reports `-2400`.

### Requirement: Keep complex picture 3-D state source-owned

The system MUST withhold the picture leaf and reject the source-bound
operation when the picture `sp3d` owner is missing, duplicated, malformed,
child-bearing, extension-bearing, or out of range.

#### Scenario: Keep a child-bearing picture owner opaque

- **WHEN** an imported picture's direct `sp3d` contains a bevel child
- **THEN** the projection emits no picture `shape3dDepthEmu` leaf
- **AND** a source-bound edit cannot target that depth as an independent
  scalar.
