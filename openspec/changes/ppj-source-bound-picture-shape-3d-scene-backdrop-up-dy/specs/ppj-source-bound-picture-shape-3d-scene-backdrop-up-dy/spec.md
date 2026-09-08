## ADDED Requirements

### Requirement: project a bounded picture backdrop up-vector Y

The presentation codec MUST project
`p:pic/p:spPr/a:scene3d/a:backdrop/a:up/@dy` through the existing
`shape3dSceneBackdropUpDyEmu` native leaf when the picture owner has exactly
one direct `a:scene3d` with exactly three direct children in
camera/light-rig/backdrop order, recognized bare camera and light rig, and one
backdrop with complete bare anchor/normal/up vectors whose signed coordinate
tokens are canonical and within the bounded range.

#### Scenario: project a source-bound picture backdrop up-vector Y

- **WHEN** a source PPTX contains a picture with a strict complete backdrop
  owner
- **THEN** the projected picture contains exactly one
  `shape3dSceneBackdropUpDyEmu` native leaf with the up-vector Y coordinate

### Requirement: edit only the picture backdrop up-vector Y token

For a projected picture backdrop-up-Y leaf, the presentation codec MUST accept
a source-bound edit only when the requested value is a changed canonical signed
coordinate token, and MUST replace only
`a:scene3d/a:backdrop/a:up/@dy` in the owning SlidePart. It MUST preserve the
anchor, normal vector, up X/Z, camera/light-rig state, picture
relationship/crop/mask/effects, other 3-D attributes, and all non-target
package parts.

#### Scenario: source-bound edit preserves picture backdrop topology

- **WHEN** a projected picture backdrop up-vector Y changes from `5` to `6`
- **THEN** only the owning SlidePart changes, the edited `dy` is `6`, and the
  picture's other XML and package parts remain intact

### Requirement: keep unsupported picture backdrops opaque

The codec MUST keep a picture backdrop-up-Y owner source-bound when the
backdrop is missing, partial, duplicated, malformed, extension-bearing, or
attached to otherwise unsupported scene topology. It MUST NOT reconstruct or
normalize such markup merely to expose or edit up-vector Y.

#### Scenario: reject an unsupported picture backdrop

- **WHEN** a picture's scene lacks a canonical complete backdrop or contains
  ambiguous scene state
- **THEN** projection does not expose `shape3dSceneBackdropUpDyEmu` for that
  owner and a source-bound edit cannot target it
