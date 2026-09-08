# ppj-source-bound-shape-3d-light-rig-rotation-longitude Specification

## ADDED Requirements

### Requirement: Strict direct light-rig rotation longitude projection

The presentation projection MUST expose
`shape3dSceneLightRigRotationLongitude60000` only when a shape has one direct
`a:scene3d` owner under `a:spPr`, and that scene has exactly one bare
`a:camera` followed by one `a:lightRig`. The camera MUST have only a
recognized `prst` attribute. The light rig MUST have only recognized `rig` and
`dir` attributes and exactly one bare `a:rot` child. The rotation MUST have
exactly `lat`, `lon`, and `rev` attributes, each a canonical non-negative
1/60000-degree token no greater than 360 degrees. Existing sibling 3-D owners
MAY remain and MUST be preserved.

#### Scenario: Project a light-rig rotation longitude

- **GIVEN** an ordinary shape with a canonical scene whose light rig has
  `rig="threePt"`, `dir="t"`, and `rot lat="0" lon="0" rev="6000000"`
- **WHEN** the PPTX is projected to PPJ
- **THEN** exactly one `shape3dSceneLightRigRotationLongitude60000` native leaf
  is emitted with value `0`

#### Scenario: Keep partial or complex rotation state source-owned

- **GIVEN** the rotation omits an angle, has an unknown/non-canonical token,
  has a duplicate/extra attribute, an extra child, or the scene has camera
  overrides/backdrop/extensions
- **WHEN** the PPTX is projected to PPJ
- **THEN** no `shape3dSceneLightRigRotationLongitude60000` leaf is emitted

### Requirement: Source-bound light-rig rotation longitude token splice

The PPJ edit plan MUST permit
`shape3dSceneLightRigRotationLongitude60000` only when the source-bound shape
still proves the strict owner and the requested value is a different canonical
non-negative rotation token no greater than 360 degrees. The PPTX edit MUST
replace only the `a:rot/@lon` token in the owning SlidePart, preserve the
scene/camera/light-rig topology and unrelated package parts, and allow the new
longitude to be recovered on a second projection.

#### Scenario: Edit and reproject light-rig rotation longitude

- **GIVEN** a projected strict scene owner with `lon="0"`
- **WHEN** PPJ changes the leaf to `1800000`
- **THEN** only the owning SlidePart changes, the package remains Open XML
  valid, the longitude token becomes `1800000`, and a second projection returns
  `1800000`

#### Scenario: Reject a non-owner edit

- **GIVEN** the source rotation is partial, child-bearing, extension-bearing,
  malformed, out of range, or the scene has unsupported state
- **WHEN** an edit targets `shape3dSceneLightRigRotationLongitude60000`
- **THEN** the edit fails closed and does not reconstruct the scene
