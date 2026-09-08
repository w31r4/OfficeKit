# Picture shape 3-D top-bevel height

## ADDED Requirements

### Requirement: Project one direct picture top-bevel height

The system MUST expose one `shape3dBevelTopHeightEmu` native leaf for an
imported picture only when its `p:spPr` contains one strict direct `a:sp3d`
owner with one direct `a:bevelT` child and a canonical non-negative `@h`
coordinate.

#### Scenario: Project a picture top-bevel height

- **WHEN** an imported picture contains direct
  `p:pic/p:spPr/a:sp3d/a:bevelT/@h="800"`
- **THEN** PPJ exposes `shape3dBevelTopHeightEmu` with value `800`
- **AND** the projection preserves the picture asset, mask, effects, and
  other direct 3-D attributes.

### Requirement: Edit only the picture top-bevel height token

The system MUST allow a source-bound `shape3dBevelTopHeightEmu` edit to
replace only the proven picture `bevelT/@h` token in the owning SlidePart.

#### Scenario: Reproject an edited picture top-bevel height

- **WHEN** the picture leaf changes from `800` to `1600`
- **THEN** only the owning SlidePart changes
- **AND** the edited package remains Open XML-valid
- **AND** a second projection reports `1600`.

### Requirement: Keep complex picture top-bevel state source-owned

The system MUST withhold the picture leaf and reject the source-bound
operation when the picture `sp3d` owner is missing, duplicated, malformed,
child-bearing beyond the strict `bevelT` profile, extension-bearing, or
out of range.

#### Scenario: Keep an unsafe picture bevel opaque

- **WHEN** an imported picture's top bevel has an unknown attribute, extra
  child, malformed height, or out-of-range height
- **THEN** the projection emits no picture `shape3dBevelTopHeightEmu` leaf
- **AND** a source-bound edit cannot target that height as an independent
  token.
