# ppj-source-bound-shape-3d-camera-preset Specification

## ADDED Requirements

### Requirement: Strict direct 3-D scene camera projection

The presentation projection MUST expose `shape3dSceneCameraPreset` only when a
shape has one direct `a:scene3d` owner under `a:spPr`, and that scene has
exactly one bare `a:camera` followed by one bare `a:lightRig`.
The camera MUST have only a recognized DrawingML preset `prst` attribute. The
light rig MUST have only recognized `rig` and `dir` attributes and neither
child may contain a child element. Existing direct `a:sp3d` scalar attributes
MAY remain and MUST be preserved.

#### Scenario: Project a bounded camera preset

- **GIVEN** an ordinary shape with a canonical two-child direct `a:scene3d` whose
  camera has `prst="orthographicFront"`
- **WHEN** the PPTX is projected to PPJ
- **THEN** exactly one `shape3dSceneCameraPreset` native leaf is emitted with
  value `orthographicFront`

#### Scenario: Keep complex scene state source-owned

- **GIVEN** the scene has camera rotation, camera FOV/zoom, a backdrop, an
  extension, an extra child, an unknown token, or a non-canonical attribute
- **WHEN** the PPTX is projected to PPJ
- **THEN** no `shape3dSceneCameraPreset` leaf is emitted

### Requirement: Source-bound camera preset token splice

The PPJ edit plan MUST permit `shape3dSceneCameraPreset` only when the
source-bound shape still proves the strict direct `a:scene3d` owner and the
requested value is a different recognized camera preset. The PPTX edit MUST replace only the
`a:camera/@prst` token in the owning SlidePart, preserve the scene/light-rig
topology and unrelated package parts, and allow the new preset to be recovered
on a second projection.

#### Scenario: Edit and reproject a camera preset

- **GIVEN** a projected strict camera owner with `orthographicFront`
- **WHEN** PPJ changes the leaf to `perspectiveFront`
- **THEN** only the owning SlidePart changes, the package remains Open XML
  valid, the camera token becomes `perspectiveFront`, and a second projection
  returns `perspectiveFront`

#### Scenario: Reject a non-owner edit

- **GIVEN** the source camera has an extra scene child, child-bearing camera,
  unknown attribute, or unsupported camera token
- **WHEN** an edit targets `shape3dSceneCameraPreset`
- **THEN** the edit fails closed and does not reconstruct the scene
