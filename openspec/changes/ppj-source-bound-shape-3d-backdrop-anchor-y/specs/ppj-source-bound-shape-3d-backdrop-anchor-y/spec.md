# ppj-source-bound-shape-3d-backdrop-anchor-y Specification

## ADDED Requirements

### Requirement: Strict direct backdrop anchor Y projection

The presentation projection MUST expose
`shape3dSceneBackdropAnchorYEmu` only when a shape has one direct `a:scene3d`
owner under `a:spPr`, followed by a bare recognized camera, a bare recognized
light rig, and a complete bare `a:backdrop`. The backdrop MUST contain exactly
one bare `a:anchor`, one bare `a:norm`, and one bare `a:up`; every coordinate
attribute MUST be present as a canonical `ST_Coordinate` literal in the
standard bounded signed EMU range. The leaf value MUST be the anchor's `y`
coordinate. Existing sibling 3-D owners MAY remain and MUST be preserved.

#### Scenario: Project a backdrop anchor Y coordinate

- **GIVEN** an ordinary shape with a canonical three-child scene whose
  backdrop has `anchor x="100" y="200" z="300"`, complete normal/up vectors,
  and whose camera/light rig are recognized bare owners
- **WHEN** the PPTX is projected to PPJ
- **THEN** exactly one `shape3dSceneBackdropAnchorYEmu` native leaf is emitted
  with value `200`

#### Scenario: Keep partial or complex backdrop source-owned

- **GIVEN** the backdrop omits a coordinate/vector, has a guide/reference or
  malformed token, has an extension/extra child, or the scene has camera/light
  transforms or other unsupported state
- **WHEN** the PPTX is projected to PPJ
- **THEN** no `shape3dSceneBackdropAnchorYEmu` leaf is emitted

### Requirement: Source-bound backdrop anchor Y token splice

The PPJ edit plan MUST permit `shape3dSceneBackdropAnchorYEmu` only when the
source-bound shape still proves the strict backdrop owner and the requested
value is a different canonical `ST_Coordinate` literal. The PPTX edit MUST
replace only `a:backdrop/a:anchor/@y` in the owning SlidePart, preserve all
other scene and package content, and allow the new coordinate to be recovered
on a second projection.

#### Scenario: Edit and reproject backdrop anchor Y

- **GIVEN** a projected strict backdrop owner with anchor `y="200"`
- **WHEN** PPJ changes the leaf to `400`
- **THEN** only the owning SlidePart changes, the package remains Open XML
  valid, the anchor token becomes `400`, and a second projection returns `400`

#### Scenario: Reject a non-owner edit

- **GIVEN** the backdrop is partial, child-bearing, extension-bearing,
  malformed, out of range, or the scene has unsupported state
- **WHEN** an edit targets `shape3dSceneBackdropAnchorYEmu`
- **THEN** the edit fails closed and does not reconstruct the scene
