## Purpose

This capability preserves and minimally edits the revolution token of a
complete 3-D camera rotation on imported pictures without owning the scene
graph.

## ADDED Requirements

### Requirement: project a bounded picture camera rotation revolution

The presentation codec MUST project
`p:pic/p:spPr/a:scene3d/a:camera/a:rot/@rev` through the existing
`shape3dSceneCameraRotationRevolution60000` native leaf when the picture owner
has exactly one direct `a:scene3d` with exactly two direct children in
camera/light-rig order, a recognized camera preset, one child-free rotation
with canonical non-negative `lat`, `lon`, and `rev` values at or below 360
degrees, and a recognized light-rig pair.

#### Scenario: project a source-bound picture rotation revolution

- **WHEN** a source PPTX contains a picture with a strict complete camera
  rotation owner
- **THEN** the projected picture contains exactly one
  `shape3dSceneCameraRotationRevolution60000` native leaf with the canonical
  numeric revolution value

### Requirement: edit only the picture camera rotation revolution token

For a projected picture camera-rotation-revolution leaf, the presentation
codec MUST accept a source-bound edit only when the requested value is a
changed canonical non-negative revolution at or below 360 degrees, and MUST
replace only `a:scene3d/a:camera/a:rot/@rev` in the owning SlidePart. It MUST
preserve the camera preset, zoom/FOV, latitude, longitude, light rig, picture
relationship/crop/mask/effects, other 3-D attributes, and all non-target
package parts.

#### Scenario: source-bound edit preserves picture rotation topology

- **WHEN** a projected picture revolution changes from `6000000` to `7200000`
- **THEN** only the owning SlidePart changes, the edited `rev` is `7200000`,
  and the picture's other XML and package parts remain intact

### Requirement: keep unsupported picture camera rotations opaque

The codec MUST keep a picture camera-rotation-revolution owner source-bound
when the rotation is missing, partial, duplicated, child-bearing, malformed,
out-of-range, extension-bearing, or attached to otherwise unsupported scene
topology. It MUST NOT reconstruct or normalize such markup merely to expose or
edit the revolution leaf.

#### Scenario: reject an unsupported picture camera rotation

- **WHEN** a picture's camera rotation owner lacks a canonical complete
  rotation or contains ambiguous scene state
- **THEN** projection does not expose
  `shape3dSceneCameraRotationRevolution60000` for that owner and a source-bound
  edit cannot target it
