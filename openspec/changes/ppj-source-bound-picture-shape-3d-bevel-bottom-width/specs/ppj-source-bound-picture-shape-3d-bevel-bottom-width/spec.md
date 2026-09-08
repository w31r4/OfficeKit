# Picture 3-D bottom-bevel width

## ADDED Requirements

### Requirement: project a bounded picture bottom-bevel width

The presentation codec MUST project
`p:pic/p:spPr/a:sp3d/a:bevelB/@w` through the existing
`shape3dBevelBottomWidthEmu` native leaf when the picture owner has exactly
one direct `a:sp3d`, exactly one direct `a:bevelB`, only the established
direct 3-D scalar attributes, only `w`, `h`, and `prst` on `bevelB`, and a
canonical non-negative signed 32-bit coordinate.

#### Scenario: project a source-bound picture bottom-bevel width

- **WHEN** a source PPTX contains a picture with a strict direct
  `a:sp3d/a:bevelB/@w` owner
- **THEN** the projected picture contains exactly one
  `shape3dBevelBottomWidthEmu` native leaf with the source coordinate

### Requirement: edit only the picture bottom-bevel width token

For a projected picture bottom-bevel width leaf, the presentation codec MUST
accept a source-bound edit only when the requested coordinate is canonical,
non-negative, within the signed 32-bit DrawingML range, and different from
the current coordinate. It MUST replace only the direct `a:bevelB/@w` lexical
value in the owning SlidePart. It MUST preserve the picture relationship,
crop, mask, effects, other direct 3-D attributes, bevel dimensions and
preset, and all non-target package parts.

#### Scenario: source-bound edit preserves picture topology

- **WHEN** a projected picture bottom-bevel width changes from `1200` to
  `2400`
- **THEN** only the owning SlidePart changes, the edited `w` is `2400`, and
  the picture's other XML and package parts remain intact

### Requirement: keep unsupported picture 3-D owners opaque

The codec MUST keep a picture bottom-bevel owner source-bound when it has an
unknown or malformed width, duplicate `a:sp3d` or `a:bevelB` state, unknown or
extension attributes, extra children, or other unsupported 3-D topology. It
MUST NOT reconstruct or normalize such markup merely to expose or edit the
width leaf.

#### Scenario: reject an unsupported picture owner

- **WHEN** a picture's 3-D owner contains unsupported or ambiguous state
- **THEN** projection does not expose `shape3dBevelBottomWidthEmu` for that
  owner and a source-bound edit cannot target it
