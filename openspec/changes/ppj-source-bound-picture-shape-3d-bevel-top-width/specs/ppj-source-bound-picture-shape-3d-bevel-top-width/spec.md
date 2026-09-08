# Picture shape 3-D top-bevel width

## ADDED Requirements

### Requirement: Project one direct picture top-bevel width

The system MUST expose one `shape3dBevelTopWidthEmu` native leaf for an
imported picture only when its `p:spPr` contains one strict direct `a:sp3d`
owner with one direct `a:bevelT` child and a canonical non-negative `@w`
coordinate.

#### Scenario: Project a picture top-bevel width

- **WHEN** an imported picture contains direct
  `p:pic/p:spPr/a:sp3d/a:bevelT/@w="1200"`
- **THEN** PPJ exposes `shape3dBevelTopWidthEmu` with value `1200`
- **AND** the projection preserves the picture asset, mask, effects, and
  other direct 3-D attributes.

### Requirement: Edit only the picture top-bevel width token

The system MUST allow a source-bound `shape3dBevelTopWidthEmu` edit to replace
only the proven picture `bevelT/@w` token in the owning SlidePart.

#### Scenario: Reproject an edited picture top-bevel width

- **WHEN** the picture leaf changes from `1200` to `2400`
- **THEN** only the owning SlidePart changes
- **AND** the edited package remains Open XML-valid
- **AND** a second projection reports `2400`.

### Requirement: Keep complex picture top-bevel state source-owned

The system MUST withhold the picture leaf and reject the source-bound
operation when the picture `sp3d` owner is missing, duplicated, malformed,
child-bearing beyond the strict `bevelT` profile, extension-bearing, or
out of range.

#### Scenario: Keep an unsafe picture bevel opaque

- **WHEN** an imported picture's top bevel has an unknown attribute, extra
  child, malformed width, or out-of-range width
- **THEN** the projection emits no picture `shape3dBevelTopWidthEmu` leaf
- **AND** a source-bound edit cannot target that width as an independent token.
