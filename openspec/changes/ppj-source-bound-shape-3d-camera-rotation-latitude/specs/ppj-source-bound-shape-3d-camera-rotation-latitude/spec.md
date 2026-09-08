# ppj-source-bound-shape-3d-camera-rotation-latitude Specification

## ADDED Requirements

### Requirement: Strict direct camera rotation latitude projection

The presentation projection MUST expose
`shape3dSceneCameraRotationLatitude60000` only when a shape has one direct
`a:scene3d` owner under `a:spPr`, and that scene has exactly one `a:camera`
followed by one `a:lightRig`. The camera MUST have only a recognized `prst`
attribute and exactly one bare `a:rot` child. The light rig MUST have only
recognized `rig` and `dir` attributes and no children. The rotation MUST have
exactly `lat`, `lon`, and `rev` attributes, each a canonical non-negative
1/60000-degree token no greater than 360 degrees. Existing sibling 3-D owners
MAY remain and MUST be preserved.

#### Scenario: Project a camera rotation latitude

- **GIVEN** an ordinary shape with a canonical scene whose camera has
  `prst="perspectiveFront"` and `rot lat="0" lon="0" rev="6000000"`, and
  whose light rig has `rig="threePt"` and `dir="t"`
- **WHEN** the PPTX is projected to PPJ
- **THEN** exactly one `shape3dSceneCameraRotationLatitude60000` native leaf
  is emitted with value `0`

#### Scenario: Keep partial or complex camera rotation source-owned

- **GIVEN** the camera rotation omits an angle, has an unknown/non-canonical
  token, has a duplicate/extra attribute, an extra child, or the scene has
  camera overrides/light rotation/backdrop/extensions
- **WHEN** the PPTX is projected to PPJ
- **THEN** no `shape3dSceneCameraRotationLatitude60000` leaf is emitted

### Requirement: Source-bound camera rotation latitude token splice

The PPJ edit plan MUST permit
`shape3dSceneCameraRotationLatitude60000` only when the source-bound shape
still proves the strict owner and the requested value is a different canonical
non-negative rotation token no greater than 360 degrees. The PPTX edit MUST
replace only the `a:rot/@lat` token in the owning SlidePart, preserve the
scene/camera/light-rig topology and unrelated package parts, and allow the new
latitude to be recovered on a second projection.

#### Scenario: Edit and reproject camera rotation latitude

- **GIVEN** a projected strict scene owner with camera `lat="0"`
- **WHEN** PPJ changes the leaf to `1800000`
- **THEN** only the owning SlidePart changes, the package remains Open XML
  valid, the latitude token becomes `1800000`, and a second projection returns
  `1800000`

#### Scenario: Reject a non-owner edit

- **GIVEN** the camera rotation is partial, child-bearing, extension-bearing,
  malformed, out of range, or the scene has unsupported state
- **WHEN** an edit targets `shape3dSceneCameraRotationLatitude60000`
- **THEN** the edit fails closed and does not reconstruct the scene
