# ppj-source-bound-shape-3d-light-rig-direction Specification

## ADDED Requirements

### Requirement: Strict direct 3-D scene light-rig direction projection

The presentation projection MUST expose
`shape3dSceneLightRigDirection` only when a shape has one direct `a:scene3d`
owner under `a:spPr`, and that scene has exactly one bare `a:camera` followed by
one bare `a:lightRig`. The camera MUST have only a recognized `prst` attribute.
The light rig MUST have only recognized `rig` and `dir` attributes and neither
child may contain a child element. Existing sibling 3-D owners MAY remain and
MUST be preserved.

#### Scenario: Project a bounded light-rig direction

- **GIVEN** an ordinary shape with a canonical two-child direct `a:scene3d`
  whose light rig has `rig="threePt"` and `dir="t"`
- **WHEN** the PPTX is projected to PPJ
- **THEN** exactly one `shape3dSceneLightRigDirection` native leaf is emitted
  with value `t`

#### Scenario: Keep complex scene state source-owned

- **GIVEN** the scene has camera rotation, light rotation, a backdrop, an
  extension, an extra child, an unknown token, or a non-canonical attribute
- **WHEN** the PPTX is projected to PPJ
- **THEN** no `shape3dSceneLightRigDirection` leaf is emitted

### Requirement: Source-bound light-rig direction token splice

The PPJ edit plan MUST permit `shape3dSceneLightRigDirection` only when the
source-bound shape still proves the strict owner and the requested value is a
different recognized light-rig direction. The PPTX edit MUST replace only the
`a:lightRig/@dir` token in the owning SlidePart, preserve the scene/camera and
light-rig preset topology and unrelated package parts, and allow the new
direction to be recovered on a second projection.

#### Scenario: Edit and reproject a light-rig direction

- **GIVEN** a projected strict scene owner with `dir="t"`
- **WHEN** PPJ changes the leaf to `br`
- **THEN** only the owning SlidePart changes, the package remains Open XML
  valid, the light-rig direction token becomes `br`, and a second projection
  returns `br`

#### Scenario: Reject a non-owner edit

- **GIVEN** the source scene has an extra child, child-bearing camera or light
  rig, unknown attribute, or unsupported light-rig direction token
- **WHEN** an edit targets `shape3dSceneLightRigDirection`
- **THEN** the edit fails closed and does not reconstruct the scene
