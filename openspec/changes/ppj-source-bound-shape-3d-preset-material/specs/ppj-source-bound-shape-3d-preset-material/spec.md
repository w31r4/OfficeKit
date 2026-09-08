# Shape 3-D preset material

## ADDED Requirements

### Requirement: Project one direct preset-material token

The system MUST expose one `shape3dPresetMaterial` native leaf only for an
ordinary shape with exactly one direct child-free `a:sp3d` whose
`prstMaterial` value is a canonical token from the bounded preset-material
vocabulary.

#### Scenario: Project a recognized material

- **WHEN** an imported shape contains a child-free direct
  `a:sp3d/@prstMaterial="metal"`
- **THEN** PPJ exposes `shape3dPresetMaterial` with value `metal`
- **AND** the projection preserves the other direct 3-D attributes as source
  state.

### Requirement: Edit only the material token

The system MUST allow a source-bound edit to replace only the proven
`prstMaterial` attribute in the owning SlidePart.

#### Scenario: Reproject an edited material

- **WHEN** the `shape3dPresetMaterial` leaf changes from `metal` to `matte`
- **THEN** only the owning SlidePart changes
- **AND** the edited package remains Open XML-valid
- **AND** a second projection reports `matte`.

### Requirement: Keep complex or unknown material state source-owned

The system MUST withhold the leaf and reject the source-bound operation when
the `sp3d` owner is missing, malformed, out of range, child-bearing,
extension-bearing, unknown, or otherwise ambiguous.

#### Scenario: Keep an unknown token opaque

- **WHEN** an imported shape's direct `sp3d` has an unrecognized
  `prstMaterial` token or a bevel child
- **THEN** the projection emits no `shape3dPresetMaterial` leaf
- **AND** a source-bound edit cannot target that material as an independent
  scalar.
