## Purpose

This capability preserves and minimally edits the backdrop normal X coordinate
of a strict 3-D scene on an imported picture without owning the scene graph.

## ADDED Requirements

### Requirement: project a bounded picture backdrop normal X

The presentation codec MUST project
`p:pic/p:spPr/a:scene3d/a:backdrop/a:norm/@dx` through the existing
`shape3dSceneBackdropNormalDxEmu` native leaf when the picture owner has
exactly one direct `a:scene3d` with exactly three direct children in
camera/light-rig/backdrop order, recognized bare camera and light rig, and one
backdrop with complete bare anchor/normal/up vectors whose signed coordinate
tokens are canonical and within the bounded range.

#### Scenario: project a source-bound picture backdrop normal X

- **WHEN** a source PPTX contains a picture with a strict complete backdrop
  owner
- **THEN** the projected picture contains exactly one
  `shape3dSceneBackdropNormalDxEmu` native leaf with the normal X coordinate

### Requirement: edit only the picture backdrop normal X token

For a projected picture backdrop-normal-X leaf, the presentation codec MUST
accept a source-bound edit only when the requested value is a changed canonical
signed coordinate token, and MUST replace only
`a:scene3d/a:backdrop/a:norm/@dx` in the owning SlidePart. It MUST preserve
the anchor, normal Y/Z, up vector, camera/light-rig state, picture
relationship/crop/mask/effects, other 3-D attributes, and all non-target
package parts.

#### Scenario: source-bound edit preserves picture backdrop topology

- **WHEN** a projected picture backdrop normal X changes from `1` to `2`
- **THEN** only the owning SlidePart changes, the edited `dx` is `2`, and the
  picture's other XML and package parts remain intact

### Requirement: keep unsupported picture backdrops opaque

The codec MUST keep a picture backdrop-normal-X owner source-bound when the
backdrop is missing, partial, duplicated, malformed, extension-bearing, or
attached to otherwise unsupported scene topology. It MUST NOT reconstruct or
normalize such markup merely to expose or edit normal X.

#### Scenario: reject an unsupported picture backdrop

- **WHEN** a picture's scene lacks a canonical complete backdrop or contains
  ambiguous scene state
- **THEN** projection does not expose `shape3dSceneBackdropNormalDxEmu` for
  that owner and a source-bound edit cannot target it
