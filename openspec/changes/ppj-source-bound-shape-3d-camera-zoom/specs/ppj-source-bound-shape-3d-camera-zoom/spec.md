# ppj-source-bound-shape-3d-camera-zoom Specification

## ADDED Requirements

### Requirement: Strict direct 3-D scene camera zoom projection

The presentation projection MUST expose
`shape3dSceneCameraZoomThousandthPercent` only when a shape has one direct
`a:scene3d` owner under `a:spPr`, and that scene has exactly one bare `a:camera`
followed by one bare `a:lightRig`. The camera MUST have exactly one recognized
`prst` and one explicit canonical non-negative `zoom` attribute. The light rig
MUST have only recognized `rig` and `dir` attributes. Neither child may contain
a child element. Existing sibling 3-D owners MAY remain and MUST be preserved.

#### Scenario: Project an explicit camera zoom override

- **GIVEN** an ordinary shape with a canonical two-child direct `a:scene3d`
  whose camera has `prst="perspectiveFront"` and `zoom="100000"`
- **WHEN** the PPTX is projected to PPJ
- **THEN** exactly one `shape3dSceneCameraZoomThousandthPercent` native leaf is
  emitted with value `100000`

#### Scenario: Keep implicit or complex camera state source-owned

- **GIVEN** the camera omits `zoom`, has an unknown/non-canonical zoom token,
  has a duplicate/extra attribute, or has a rotation/FOV override
- **WHEN** the PPTX is projected to PPJ
- **THEN** no `shape3dSceneCameraZoomThousandthPercent` leaf is emitted

### Requirement: Source-bound camera zoom token splice

The PPJ edit plan MUST permit `shape3dSceneCameraZoomThousandthPercent` only
when the source-bound shape still proves the strict owner and the requested
value is a different canonical non-negative zoom token in the bounded native
range. The PPTX edit MUST replace only the `a:camera/@zoom` token in the owning
SlidePart, preserve the scene/light-rig topology and unrelated package parts,
and allow the new zoom to be recovered on a second projection.

#### Scenario: Edit and reproject camera zoom

- **GIVEN** a projected strict scene owner with `zoom="100000"`
- **WHEN** PPJ changes the leaf to `50000`
- **THEN** only the owning SlidePart changes, the package remains Open XML
  valid, the camera zoom token becomes `50000`, and a second projection returns
  `50000`

#### Scenario: Reject a non-owner edit

- **GIVEN** the source camera omits zoom or has rotation, FOV, an extra child,
  unknown attribute, malformed token, or unsupported scene state
- **WHEN** an edit targets `shape3dSceneCameraZoomThousandthPercent`
- **THEN** the edit fails closed and does not reconstruct the scene
