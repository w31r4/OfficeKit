# Picture 3-D contour RGB

## ADDED Requirements

### Requirement: project a bounded picture 3-D contour color

The presentation codec MUST project
`p:pic/p:spPr/a:sp3d/a:contourClr/a:srgbClr/@val` through the existing
`shape3dContourRgb` native leaf when the picture owner has exactly one
direct `a:sp3d`, exactly one direct `a:contourClr`, exactly one direct
`a:srgbClr`, only the established direct 3-D scalar attributes, no
attributes or extra children on the color owner, and a bare six-digit RGB
token.

#### Scenario: project a source-bound picture contour color

- **WHEN** a source PPTX contains a picture with a strict direct
  `a:sp3d/a:contourClr/a:srgbClr/@val` owner
- **THEN** the projected picture contains exactly one `shape3dContourRgb`
  native leaf with the canonical `#rrggbb` value

### Requirement: edit only the picture contour color token

For a projected picture contour-color leaf, the presentation codec MUST
accept a source-bound edit only when the requested value is a canonical,
changed RGB token, and MUST replace only the direct
`a:sp3d/a:contourClr/a:srgbClr/@val` lexical value in the owning SlidePart.
It MUST preserve the picture relationship, crop, mask, effects, other direct
3-D attributes, and all non-target package parts.

#### Scenario: source-bound edit preserves picture topology

- **WHEN** a projected picture contour color changes from `#336699` to
  `#6699CC`
- **THEN** only the owning SlidePart changes, the edited `val` is `6699CC`,
  and the picture's other XML and package parts remain intact

### Requirement: keep unsupported picture 3-D colors opaque

The codec MUST keep a picture contour-color owner source-bound when it has an
unknown or malformed color token, duplicate `a:sp3d` or color-owner state,
unknown or extension attributes, extra children, or other unsupported 3-D
topology. It MUST NOT reconstruct or normalize such markup merely to expose
or edit the color leaf.

#### Scenario: reject an unsupported picture color owner

- **WHEN** a picture's 3-D contour-color owner contains unsupported or
  ambiguous state
- **THEN** projection does not expose `shape3dContourRgb` for that owner and
  a source-bound edit cannot target it
