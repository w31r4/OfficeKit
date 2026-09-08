# ppj-source-bound-shape-3d-light-rig-preset Specification

## ADDED Requirements

### Requirement: Strict direct 3-D scene light-rig projection

The presentation projection MUST expose `shape3dSceneLightRigPreset` only when
a shape has one direct `a:scene3d` owner under `a:spPr`, and that scene has
exactly one bare `a:camera` followed by one bare `a:lightRig`. The camera MUST
have only a recognized `prst` attribute. The light rig MUST have only
recognized `rig` and `dir` attributes and neither child may contain a child
element. Existing sibling 3-D owners MAY remain and MUST be preserved.

#### Scenario: Project a bounded light-rig preset

- **GIVEN** an ordinary shape with a canonical two-child direct `a:scene3d`
  whose light rig has `rig="threePt"` and `dir="t"`
- **WHEN** the PPTX is projected to PPJ
- **THEN** exactly one `shape3dSceneLightRigPreset` native leaf is emitted with
  value `threePt`

#### Scenario: Keep complex scene state source-owned

- **GIVEN** the scene has camera rotation, light rotation, a backdrop, an
  extension, an extra child, an unknown token, or a non-canonical attribute
- **WHEN** the PPTX is projected to PPJ
- **THEN** no `shape3dSceneLightRigPreset` leaf is emitted

### Requirement: Source-bound light-rig preset token splice

The PPJ edit plan MUST permit `shape3dSceneLightRigPreset` only when the
source-bound shape still proves the strict owner and the requested value is a
different recognized light-rig preset. The PPTX edit MUST replace only the
`a:lightRig/@rig` token in the owning SlidePart, preserve the scene/camera and
light direction topology and unrelated package parts, and allow the new preset
to be recovered on a second projection.

#### Scenario: Edit and reproject a light-rig preset

- **GIVEN** a projected strict scene owner with `threePt`
- **WHEN** PPJ changes the leaf to `soft`
- **THEN** only the owning SlidePart changes, the package remains Open XML
  valid, the light-rig token becomes `soft`, and a second projection returns
  `soft`

#### Scenario: Reject a non-owner edit

- **GIVEN** the source scene has an extra child, child-bearing camera or light
  rig, unknown attribute, or unsupported light-rig token
- **WHEN** an edit targets `shape3dSceneLightRigPreset`
- **THEN** the edit fails closed and does not reconstruct the scene
