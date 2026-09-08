## Purpose

This capability preserves and minimally edits an explicit 3-D camera zoom
token on imported pictures without taking ownership of the surrounding scene.

## ADDED Requirements

### Requirement: project a bounded picture 3-D scene camera zoom

The presentation codec MUST project
`p:pic/p:spPr/a:scene3d/a:camera/@zoom` through the existing
`shape3dSceneCameraZoomThousandthPercent` native leaf when the picture owner
has exactly one direct `a:scene3d` with exactly two direct children in
camera/light-rig order, a recognized camera preset with a canonical explicit
non-negative zoom, and a recognized light-rig pair.

#### Scenario: project a source-bound picture camera zoom

- **WHEN** a source PPTX contains a picture with a strict direct
  `a:scene3d/a:camera/@zoom` owner
- **THEN** the projected picture contains exactly one
  `shape3dSceneCameraZoomThousandthPercent` native leaf with the canonical
  numeric zoom value

### Requirement: edit only the picture camera zoom token

For a projected picture camera-zoom leaf, the presentation codec MUST accept a
source-bound edit only when the requested value is a changed canonical
non-negative zoom token, and MUST replace only the direct
`a:scene3d/a:camera/@zoom` lexical value in the owning SlidePart. It MUST
preserve the picture relationship, crop, mask, effects, camera preset, light
rig, other 3-D attributes, and all non-target package parts.

#### Scenario: source-bound edit preserves picture topology

- **WHEN** a projected picture camera zoom changes from `100000` to `50000`
- **THEN** only the owning SlidePart changes, the edited `zoom` is `50000`,
  and the picture's other XML and package parts remain intact

### Requirement: keep unsupported picture camera zoom owners opaque

The codec MUST keep a picture camera-zoom owner source-bound when zoom is
omitted/default, negative, malformed, duplicated, extension-bearing, or
attached to an unknown, partial, transformed, or otherwise unsupported scene
topology. It MUST NOT reconstruct or normalize such markup merely to expose or
edit the camera zoom leaf.

#### Scenario: reject an unsupported picture camera zoom owner

- **WHEN** a picture's camera zoom owner contains omitted/default or ambiguous
  scene state
- **THEN** projection does not expose
  `shape3dSceneCameraZoomThousandthPercent` for that owner and a source-bound
  edit cannot target it
