## Purpose

This capability preserves and minimally edits the latitude token of a complete
3-D camera rotation on imported pictures without owning the scene graph.

## ADDED Requirements

### Requirement: project a bounded picture camera rotation latitude

The presentation codec MUST project
`p:pic/p:spPr/a:scene3d/a:camera/a:rot/@lat` through the existing
`shape3dSceneCameraRotationLatitude60000` native leaf when the picture owner
has exactly one direct `a:scene3d` with exactly two direct children in
camera/light-rig order, a recognized camera preset, one child-free rotation
with canonical non-negative `lat`, `lon`, and `rev` values at or below 360
degrees, and a recognized light-rig pair.

#### Scenario: project a source-bound picture rotation latitude

- **WHEN** a source PPTX contains a picture with a strict complete camera
  rotation owner
- **THEN** the projected picture contains exactly one
  `shape3dSceneCameraRotationLatitude60000` native leaf with the canonical
  numeric latitude value

### Requirement: edit only the picture camera rotation latitude token

For a projected picture camera-rotation-latitude leaf, the presentation codec
MUST accept a source-bound edit only when the requested value is a changed
canonical non-negative latitude at or below 360 degrees, and MUST replace only
`a:scene3d/a:camera/a:rot/@lat` in the owning SlidePart. It MUST preserve the
camera preset, zoom/FOV, longitude, revolution, light rig, picture
relationship/crop/mask/effects, other 3-D attributes, and all non-target
package parts.

#### Scenario: source-bound edit preserves picture rotation topology

- **WHEN** a projected picture latitude changes from `0` to `1800000`
- **THEN** only the owning SlidePart changes, the edited `lat` is `1800000`,
  and the picture's other XML and package parts remain intact

### Requirement: keep unsupported picture camera rotations opaque

The codec MUST keep a picture camera-rotation-latitude owner source-bound when
the rotation is missing, partial, duplicated, child-bearing, malformed,
out-of-range, extension-bearing, or attached to otherwise unsupported scene
topology. It MUST NOT reconstruct or normalize such markup merely to expose or
edit the latitude leaf.

#### Scenario: reject an unsupported picture camera rotation

- **WHEN** a picture's camera rotation owner lacks a canonical complete
  rotation or contains ambiguous scene state
- **THEN** projection does not expose
  `shape3dSceneCameraRotationLatitude60000` for that owner and a source-bound
  edit cannot target it
