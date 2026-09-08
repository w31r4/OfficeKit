# ppj-source-bound-shape-3d-camera-fov Specification

## ADDED Requirements

### Requirement: Strict direct 3-D scene camera FOV projection

The presentation projection MUST expose `shape3dSceneCameraFov60000` only when
a shape has one direct `a:scene3d` owner under `a:spPr`, and that scene has
exactly one bare `a:camera` followed by one bare `a:lightRig`. The camera MUST
have exactly one recognized `prst` and one explicit canonical FOV attribute in
the strict positive-below-180-degree range; an explicit canonical `zoom` may
also be present. The light rig MUST have only recognized `rig` and `dir`
attributes. Neither child may contain a child element. Existing sibling 3-D
owners MAY remain and MUST be preserved.

#### Scenario: Project an explicit camera FOV override

- **GIVEN** an ordinary shape with a canonical two-child direct `a:scene3d`
  whose camera has `prst="perspectiveFront"` and `fov="2700000"`
- **WHEN** the PPTX is projected to PPJ
- **THEN** exactly one `shape3dSceneCameraFov60000` native leaf is emitted with
  value `2700000`

#### Scenario: Keep implicit or complex camera state source-owned

- **GIVEN** the camera omits `fov`, has an out-of-range/non-canonical FOV token,
  has a duplicate/extra attribute, or has a rotation override
- **WHEN** the PPTX is projected to PPJ
- **THEN** no `shape3dSceneCameraFov60000` leaf is emitted

### Requirement: Source-bound camera FOV token splice

The PPJ edit plan MUST permit `shape3dSceneCameraFov60000` only when the
source-bound shape still proves the strict owner and the requested value is a
different canonical FOV token in the bounded native range. The PPTX edit MUST
replace only the `a:camera/@fov` token in the owning SlidePart, preserve the
scene/zoom/light-rig topology and unrelated package parts, and allow the new
FOV to be recovered on a second projection.

#### Scenario: Edit and reproject camera FOV

- **GIVEN** a projected strict scene owner with `fov="2700000"`
- **WHEN** PPJ changes the leaf to `5400000`
- **THEN** only the owning SlidePart changes, the package remains Open XML
  valid, the camera FOV token becomes `5400000`, and a second projection
  returns `5400000`

#### Scenario: Reject a non-owner edit

- **GIVEN** the source camera omits FOV or has rotation, an extra child,
  unknown attribute, malformed token, or unsupported scene state
- **WHEN** an edit targets `shape3dSceneCameraFov60000`
- **THEN** the edit fails closed and does not reconstruct the scene
