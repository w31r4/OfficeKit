## Purpose

Expose one directly provable 3-D contour RGB token so PPJ can edit that small
part of an imported PowerPoint shape without flattening its surrounding scene.

## ADDED Requirements

### Requirement: Project one direct contour RGB token

The system MUST expose one `shape3dContourRgb` native leaf only for an ordinary
shape with exactly one direct `a:sp3d` whose only child is `a:contourClr`, whose
only color child is a bare `a:srgbClr`, and whose `val` is a six-digit RGB token.

#### Scenario: Project a recognized contour color

- **WHEN** an imported shape contains direct
  `a:sp3d/a:contourClr/a:srgbClr/@val="336699"` with bounded root attributes
- **THEN** PPJ exposes `shape3dContourRgb` with value `#336699`
- **AND** the projection retains the source-bound 3-D owner.

### Requirement: Edit only the contour RGB token

The system MUST allow a source-bound edit to replace only the proven
`a:srgbClr/@val` token in the owning SlidePart.

#### Scenario: Reproject an edited contour color

- **WHEN** `shape3dContourRgb` changes from `#336699` to `#6699CC`
- **THEN** only the owning SlidePart changes
- **AND** the edited package remains Open XML-valid
- **AND** a second projection reports `#6699CC`.

### Requirement: Keep complex contour color source-owned

The system MUST withhold the leaf and reject a source-bound operation when the
`sp3d` owner is missing, duplicated, has another 3-D child, has extension or
unknown state, or the contour color contains transforms, children, a non-RGB
color model, a missing value, or a noncanonical RGB token.

#### Scenario: Keep a transformed contour color opaque

- **WHEN** `a:contourClr/a:srgbClr` contains an alpha or other transform child
- **THEN** the projection emits no `shape3dContourRgb` leaf
- **AND** a source-bound edit cannot target that color as an independent scalar.
