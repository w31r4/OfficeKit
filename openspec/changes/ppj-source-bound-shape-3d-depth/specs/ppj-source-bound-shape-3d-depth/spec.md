# Shape 3-D depth

## ADDED Requirements

### Requirement: Project one direct shape depth scalar

The system MUST expose one `shape3dDepthEmu` native leaf only for an ordinary
shape with exactly one direct child-free `a:sp3d` whose `z` value is a
canonical signed 32-bit integer.

#### Scenario: Project a negative direct depth

- **WHEN** an imported shape contains a child-free direct `a:sp3d/@z="-1200"`
- **THEN** PPJ exposes `shape3dDepthEmu` with value `-1200`
- **AND** the projection preserves the other direct 3-D attributes as source
  state.

### Requirement: Edit only the depth token

The system MUST allow a source-bound edit to replace only the proven `z`
attribute in the owning SlidePart.

#### Scenario: Reproject an edited depth

- **WHEN** the `shape3dDepthEmu` leaf changes from `-1200` to `-2400`
- **THEN** only the owning SlidePart changes
- **AND** the edited package remains Open XML-valid
- **AND** a second projection reports `-2400`.

### Requirement: Keep complex 3-D state source-owned

The system MUST withhold the leaf and reject the source-bound operation when
the `sp3d` owner is missing, malformed, out of range, child-bearing,
extension-bearing, or otherwise ambiguous.

#### Scenario: Keep a child-bearing owner opaque

- **WHEN** an imported shape's direct `sp3d` contains a bevel child or an
  extension-bearing attribute set
- **THEN** the projection emits no `shape3dDepthEmu` leaf
- **AND** a source-bound edit cannot target that depth as an independent
  scalar.
