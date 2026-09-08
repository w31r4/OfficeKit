# ppj-source-bound-shape-3d-backdrop-up-dy Specification

## ADDED Requirements

### Requirement: Strict direct backdrop up-vector Y projection

The presentation projection MUST expose `shape3dSceneBackdropUpDyEmu` only
when a shape has one direct `a:scene3d` owner under `a:spPr`, followed by a
bare recognized camera, a bare recognized light rig, and a complete bare
`a:backdrop`. The backdrop MUST contain exactly one bare `a:anchor`, one bare
`a:norm`, and one bare `a:up`; every coordinate attribute MUST be present as a
canonical `ST_Coordinate` literal in the standard bounded signed EMU range.
The leaf value MUST be the up vector's `dy` coordinate. Existing sibling 3-D
owners MAY remain and MUST be preserved.

#### Scenario: Project a backdrop up-vector Y coordinate

- **GIVEN** an ordinary shape with a canonical three-child scene whose
  backdrop has complete anchor, normal, and `up dx="0" dy="1" dz="0"` vectors,
  and whose camera/light rig are recognized bare owners
- **WHEN** the PPTX is projected to PPJ
- **THEN** exactly one `shape3dSceneBackdropUpDyEmu` native leaf is emitted
  with value `1`

#### Scenario: Keep partial or complex backdrop source-owned

- **GIVEN** the backdrop omits a coordinate/vector, has a guide/reference or
  malformed token, has an extension/extra child, or the scene has camera/light
  transforms or other unsupported state
- **WHEN** the PPTX is projected to PPJ
- **THEN** no `shape3dSceneBackdropUpDyEmu` leaf is emitted

### Requirement: Source-bound backdrop up-vector Y token splice

The PPJ edit plan MUST permit `shape3dSceneBackdropUpDyEmu` only when the
source-bound shape still proves the strict backdrop owner and the requested
value is a different canonical `ST_Coordinate` literal. The PPTX edit MUST
replace only `a:backdrop/a:up/@dy` in the owning SlidePart, preserve all other
scene and package content, and allow the new coordinate to be recovered on a
second projection.

#### Scenario: Edit and reproject backdrop up-vector Y

- **GIVEN** a projected strict backdrop owner with `up dy="1"`
- **WHEN** PPJ changes the leaf to `500`
- **THEN** only the owning SlidePart changes, the package remains Open XML
  valid, the up-vector token becomes `500`, and a second projection returns
  `500`

#### Scenario: Reject a non-owner edit

- **GIVEN** the backdrop is partial, child-bearing, extension-bearing,
  malformed, out of range, or the scene has unsupported state
- **WHEN** an edit targets `shape3dSceneBackdropUpDyEmu`
- **THEN** the edit fails closed and does not reconstruct the scene
