## Purpose

This capability preserves and minimally edits an explicit 3-D camera FOV token
on imported pictures while leaving the surrounding scene source-owned.

## ADDED Requirements

### Requirement: project a bounded picture 3-D scene camera FOV

The presentation codec MUST project
`p:pic/p:spPr/a:scene3d/a:camera/@fov` through the existing
`shape3dSceneCameraFov60000` native leaf when the picture owner has exactly one
direct `a:scene3d` with exactly two direct children in camera/light-rig order,
a recognized camera preset with a canonical positive FOV below 180 degrees,
and a recognized light-rig pair.

#### Scenario: project a source-bound picture camera FOV

- **WHEN** a source PPTX contains a picture with a strict direct
  `a:scene3d/a:camera/@fov` owner
- **THEN** the projected picture contains exactly one
  `shape3dSceneCameraFov60000` native leaf with the canonical numeric FOV
  value

### Requirement: edit only the picture camera FOV token

For a projected picture camera-FOV leaf, the presentation codec MUST accept a
source-bound edit only when the requested value is a changed canonical positive
FOV below 180 degrees, and MUST replace only the direct
`a:scene3d/a:camera/@fov` lexical value in the owning SlidePart. It MUST
preserve the picture relationship, crop, mask, effects, camera preset, zoom,
light rig, other 3-D attributes, and all non-target package parts.

#### Scenario: source-bound edit preserves picture topology

- **WHEN** a projected picture camera FOV changes from `2700000` to `5400000`
- **THEN** only the owning SlidePart changes, the edited `fov` is `5400000`,
  and the picture's other XML and package parts remain intact

### Requirement: keep unsupported picture camera FOV owners opaque

The codec MUST keep a picture camera-FOV owner source-bound when FOV is
omitted, non-positive, at or above 180 degrees, malformed, duplicated,
extension-bearing, or attached to an unknown, partial, transformed, or
otherwise unsupported scene topology. It MUST NOT reconstruct or normalize such
markup merely to expose or edit the camera FOV leaf.

#### Scenario: reject an unsupported picture camera FOV owner

- **WHEN** a picture's camera FOV owner contains omitted, invalid, or ambiguous
  scene state
- **THEN** projection does not expose `shape3dSceneCameraFov60000` for that
  owner and a source-bound edit cannot target it
