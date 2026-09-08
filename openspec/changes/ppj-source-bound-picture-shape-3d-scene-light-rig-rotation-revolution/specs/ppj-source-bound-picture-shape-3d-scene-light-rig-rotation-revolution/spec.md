## Purpose

This capability preserves and minimally edits the light-rig rotation
revolution of a strict 3-D scene on an imported picture without owning the
scene graph.

## ADDED Requirements

### Requirement: project a bounded picture light-rig rotation revolution

The presentation codec MUST project
`p:pic/p:spPr/a:scene3d/a:lightRig/a:rot/@rev` through the existing
`shape3dSceneLightRigRotationRevolution60000` native leaf when the picture
owner has exactly one direct `a:scene3d` with exactly two direct children in
camera/light-rig order, a recognized bare camera, a recognized bare light rig,
and exactly one child-free rotation carrying canonical non-negative
`lat`/`lon`/`rev` values at or below 360 degrees.

#### Scenario: project a source-bound picture light-rig revolution

- **WHEN** a source PPTX contains a picture with a strict complete light-rig
  rotation owner
- **THEN** the projected picture contains exactly one
  `shape3dSceneLightRigRotationRevolution60000` native leaf with the revolution
  in 60000ths-of-a-degree units

### Requirement: edit only the picture light-rig revolution token

For a projected picture light-rig-revolution leaf, the presentation codec MUST
accept a source-bound edit only when the requested value is a changed canonical
angle token, and MUST replace only `a:scene3d/a:lightRig/a:rot/@rev` in the
owning SlidePart. It MUST preserve latitude, longitude, camera/light-rig
state, picture relationship/crop/mask/effects, other 3-D attributes, and all
non-target package parts.

#### Scenario: source-bound edit preserves picture rotation topology

- **WHEN** a projected picture light-rig revolution changes from `6000000` to
  `7200000`
- **THEN** only the owning SlidePart changes, the edited `rev` is `7200000`,
  and the picture's other XML and package parts remain intact

### Requirement: keep unsupported picture light-rig rotations opaque

The codec MUST keep a picture light-rig-revolution owner source-bound when the
camera/light-rig rotation is missing, partial, duplicated, malformed,
extension-bearing, or attached to otherwise unsupported scene topology. It
MUST NOT reconstruct or normalize such markup merely to expose or edit the
revolution.

#### Scenario: reject an unsupported picture light-rig rotation

- **WHEN** a picture's scene lacks a canonical complete rotation or contains
  ambiguous scene state
- **THEN** projection does not expose
  `shape3dSceneLightRigRotationRevolution60000` for that owner and a
  source-bound edit cannot target it
