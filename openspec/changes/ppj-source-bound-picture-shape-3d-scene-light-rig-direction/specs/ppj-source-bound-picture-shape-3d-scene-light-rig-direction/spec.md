## Purpose

This capability preserves and minimally edits the light-rig direction of a
strict 3-D scene on an imported picture without owning the scene graph.

## ADDED Requirements

### Requirement: project a bounded picture light-rig direction

The presentation codec MUST project
`p:pic/p:spPr/a:scene3d/a:lightRig/@dir` through the existing
`shape3dSceneLightRigDirection` native leaf when the picture owner has exactly
one direct `a:scene3d` with exactly two direct children in camera/light-rig
order, a recognized camera preset, a child-free light rig with canonical
recognized `rig` and `dir` attributes, and no unknown scene attributes or
children.

#### Scenario: project a source-bound picture light-rig direction

- **WHEN** a source PPTX contains a picture with a strict camera/light-rig
  pair
- **THEN** the projected picture contains exactly one
  `shape3dSceneLightRigDirection` native leaf with the canonical direction
  token

### Requirement: edit only the picture light-rig direction token

For a projected picture light-rig-direction leaf, the presentation codec MUST
accept a source-bound edit only when the requested value is a changed
recognized light-rig direction token, and MUST replace only
`a:scene3d/a:lightRig/@dir` in the owning SlidePart. It MUST preserve the
camera, light-rig preset, picture relationship/crop/mask/effects, other 3-D
attributes, and all non-target package parts.

#### Scenario: source-bound edit preserves picture scene topology

- **WHEN** a projected picture light-rig direction changes from `t` to `br`
- **THEN** only the owning SlidePart changes, the edited `dir` is `br`, and the
  picture's other XML and package parts remain intact

### Requirement: keep unsupported picture light-rig owners opaque

The codec MUST keep a picture light-rig-direction owner source-bound when the
camera/light-rig pair is missing, partial, duplicated, malformed,
extension-bearing, or attached to otherwise unsupported scene topology. It
MUST NOT reconstruct or normalize such markup merely to expose or edit the
light-rig direction.

#### Scenario: reject an unsupported picture light-rig owner

- **WHEN** a picture's scene lacks a canonical camera/light-rig pair or
  contains ambiguous scene state
- **THEN** projection does not expose `shape3dSceneLightRigDirection` for that
  owner and a source-bound edit cannot target it
