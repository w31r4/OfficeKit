## ADDED Requirements

### Requirement: project a bounded 3-D extrusion theme token

The presentation codec MUST project a direct `a:schemeClr/@val` under
`p:sp/p:spPr/a:sp3d/a:extrusionClr` or
`p:pic/p:spPr/a:sp3d/a:extrusionClr` through the
`shape3dExtrusionColorScheme` native leaf when the 3-D root and color owner
have only the bounded recognized attributes and the scheme color is child-free.

#### Scenario: project shape and picture extrusion theme colors

- **WHEN** a source PPTX contains a shape or picture with a strict direct 3-D
  extrusion `schemeClr` owner
- **THEN** the projected element contains one
  `shape3dExtrusionColorScheme` native leaf with the canonical theme token

### Requirement: edit only the extrusion theme token

For a projected extrusion-theme leaf, the presentation codec MUST accept a
source-bound edit only when the requested supported theme token differs from
the expected token, and MUST replace only
`a:sp3d/a:extrusionClr/a:schemeClr/@val` in the owning SlidePart. It MUST
preserve all other 3-D state, picture relationships/crop/mask/effects, and
non-target package parts.

#### Scenario: source-bound extrusion theme edit preserves the owner

- **WHEN** a projected shape or picture extrusion token changes from `accent1`
  to `accent2`
- **THEN** only the owning SlidePart changes, the edited `schemeClr/@val` is
  `accent2`, and all non-target state remains intact

### Requirement: keep unsupported extrusion color graphs opaque

The codec MUST keep an extrusion color owner source-bound when the color child
has transforms, extra attributes/children, duplicate owners, unknown scheme
values, or otherwise ambiguous 3-D topology. It MUST NOT normalize such markup
merely to expose or edit the theme token.

#### Scenario: reject transformed extrusion theme colors

- **WHEN** a shape or picture extrusion `schemeClr` contains an unsupported
  transform or extension
- **THEN** projection does not expose `shape3dExtrusionColorScheme` for that
  owner and a source-bound edit cannot target it
