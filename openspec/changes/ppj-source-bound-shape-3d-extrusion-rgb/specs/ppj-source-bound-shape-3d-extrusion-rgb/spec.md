## Purpose

Expose one directly provable 3-D extrusion RGB token so PPJ can edit that
small part of an imported PowerPoint shape without flattening its scene.

## ADDED Requirements

### Requirement: Project one direct extrusion RGB token

The system MUST expose one `shape3dExtrusionRgb` native leaf only for an
ordinary shape with exactly one direct `a:sp3d` whose only child is
`a:extrusionClr`, whose only color child is a bare `a:srgbClr`, and whose
`val` is a six-digit RGB token.

#### Scenario: Project a recognized extrusion color

- **WHEN** an imported shape contains direct
  `a:sp3d/a:extrusionClr/a:srgbClr/@val="663399"` with bounded root attributes
- **THEN** PPJ exposes `shape3dExtrusionRgb` with value `#663399`
- **AND** the projection retains the source-bound 3-D owner.

### Requirement: Edit only the extrusion RGB token

The system MUST allow a source-bound edit to replace only the proven
`a:srgbClr/@val` token in the owning SlidePart.

#### Scenario: Reproject an edited extrusion color

- **WHEN** `shape3dExtrusionRgb` changes from `#663399` to `#9966CC`
- **THEN** only the owning SlidePart changes
- **AND** the edited package remains Open XML-valid
- **AND** a second projection reports `#9966CC`.

### Requirement: Keep complex extrusion color source-owned

The system MUST withhold the leaf and reject a source-bound operation when the
`sp3d` owner is missing, duplicated, has another 3-D child, has extension or
unknown state, or the extrusion color contains transforms, children, a non-RGB
color model, a missing value, or a noncanonical RGB token.

#### Scenario: Keep a transformed extrusion color opaque

- **WHEN** `a:extrusionClr/a:srgbClr` contains an alpha or other transform child
- **THEN** the projection emits no `shape3dExtrusionRgb` leaf
- **AND** a source-bound edit cannot target that color as an independent scalar.
