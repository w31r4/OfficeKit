# Picture shape 3-D extrusion height

## ADDED Requirements

### Requirement: Project one direct picture extrusion-height scalar

The system MUST expose one `shape3dExtrusionHeightEmu` native leaf for an
imported picture only when its `p:spPr` contains one child-free direct
`a:sp3d` with a canonical non-negative signed 32-bit `extrusionH` coordinate.

#### Scenario: Project a picture extrusion height

- **WHEN** an imported picture contains direct
  `p:pic/p:spPr/a:sp3d/@extrusionH="1200"`
- **THEN** PPJ exposes `shape3dExtrusionHeightEmu` with value `1200`
- **AND** the projection preserves the picture asset, mask, effects, and
  other direct 3-D attributes as source state.

### Requirement: Edit only the picture extrusion-height token

The system MUST allow a source-bound `shape3dExtrusionHeightEmu` edit to
replace only the proven picture `sp3d/@extrusionH` token in the owning
SlidePart.

#### Scenario: Reproject an edited picture extrusion height

- **WHEN** the picture leaf changes from `1200` to `2400`
- **THEN** only the owning SlidePart changes
- **AND** the edited package remains Open XML-valid
- **AND** a second projection reports `2400`.

### Requirement: Keep complex picture 3-D state source-owned

The system MUST withhold the picture leaf and reject the source-bound
operation when the picture `sp3d` owner is missing, duplicated, malformed,
child-bearing, extension-bearing, or out of range.

#### Scenario: Keep a child-bearing picture owner opaque

- **WHEN** an imported picture's direct `sp3d` contains a bevel child
- **THEN** the projection emits no picture `shape3dExtrusionHeightEmu` leaf
- **AND** a source-bound edit cannot target that extrusion height as an
  independent scalar.
