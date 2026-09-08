# Picture shape 3-D preset material

## ADDED Requirements

### Requirement: Project one direct picture preset-material token

The system MUST expose one `shape3dPresetMaterial` native leaf for an
imported picture only when its `p:spPr` contains one child-free direct
`a:sp3d` with a canonical token from the bounded preset-material vocabulary.

#### Scenario: Project a picture preset material

- **WHEN** an imported picture contains direct
  `p:pic/p:spPr/a:sp3d/@prstMaterial="metal"`
- **THEN** PPJ exposes `shape3dPresetMaterial` with value `metal`
- **AND** the projection preserves the picture asset, mask, effects, and
  other direct 3-D attributes as source state.

### Requirement: Edit only the picture preset-material token

The system MUST allow a source-bound `shape3dPresetMaterial` edit to replace
only the proven picture `sp3d/@prstMaterial` token in the owning SlidePart.

#### Scenario: Reproject an edited picture preset material

- **WHEN** the picture leaf changes from `metal` to `matte`
- **THEN** only the owning SlidePart changes
- **AND** the edited package remains Open XML-valid
- **AND** a second projection reports `matte`.

### Requirement: Keep complex picture 3-D state source-owned

The system MUST withhold the picture leaf and reject the source-bound
operation when the picture `sp3d` owner is missing, duplicated, malformed,
child-bearing, extension-bearing, unknown, or out of range.

#### Scenario: Keep an unknown picture material opaque

- **WHEN** an imported picture's direct `sp3d/@prstMaterial` is outside the
  bounded vocabulary
- **THEN** the projection emits no picture `shape3dPresetMaterial` leaf
- **AND** a source-bound edit cannot target that material as an independent
  token.
