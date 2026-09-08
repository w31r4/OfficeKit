## ADDED Requirements

### Requirement: project a bounded picture backdrop up-vector X

The presentation codec MUST project
`p:pic/p:spPr/a:scene3d/a:backdrop/a:up/@dx` through the existing
`shape3dSceneBackdropUpDxEmu` native leaf when the picture owner has exactly
one direct `a:scene3d` with exactly three direct children in
camera/light-rig/backdrop order, recognized bare camera and light rig, and one
backdrop with complete bare anchor/normal/up vectors whose signed coordinate
tokens are canonical and within the bounded range.

#### Scenario: project a source-bound picture backdrop up-vector X

- **WHEN** a source PPTX contains a picture with a strict complete backdrop
  owner
- **THEN** the projected picture contains exactly one
  `shape3dSceneBackdropUpDxEmu` native leaf with the up-vector X coordinate

### Requirement: edit only the picture backdrop up-vector X token

For a projected picture backdrop-up-X leaf, the presentation codec MUST accept
a source-bound edit only when the requested value is a changed canonical signed
coordinate token, and MUST replace only
`a:scene3d/a:backdrop/a:up/@dx` in the owning SlidePart. It MUST preserve the
anchor, normal vector, up Y/Z, camera/light-rig state, picture
relationship/crop/mask/effects, other 3-D attributes, and all non-target
package parts.

#### Scenario: source-bound edit preserves picture backdrop topology

- **WHEN** a projected picture backdrop up-vector X changes from `4` to `5`
- **THEN** only the owning SlidePart changes, the edited `dx` is `5`, and the
  picture's other XML and package parts remain intact

### Requirement: keep unsupported picture backdrops opaque

The codec MUST keep a picture backdrop-up-X owner source-bound when the
backdrop is missing, partial, duplicated, malformed, extension-bearing, or
attached to otherwise unsupported scene topology. It MUST NOT reconstruct or
normalize such markup merely to expose or edit up-vector X.

#### Scenario: reject an unsupported picture backdrop

- **WHEN** a picture's scene lacks a canonical complete backdrop or contains
  ambiguous scene state
- **THEN** projection does not expose `shape3dSceneBackdropUpDxEmu` for that
  owner and a source-bound edit cannot target it
