# Picture 3-D bottom-bevel height

## ADDED Requirements

### Requirement: project a bounded picture bottom-bevel height

The presentation codec MUST project
`p:pic/p:spPr/a:sp3d/a:bevelB/@h` through the existing
`shape3dBevelBottomHeightEmu` native leaf when the picture owner has exactly
one direct `a:sp3d`, exactly one direct `a:bevelB`, only the established
direct 3-D scalar attributes, only `w`, `h`, and `prst` on `bevelB`, and a
canonical non-negative signed 32-bit coordinate.

#### Scenario: project a source-bound picture bottom-bevel height

- **WHEN** a source PPTX contains a picture with a strict direct
  `a:sp3d/a:bevelB/@h` owner
- **THEN** the projected picture contains exactly one
  `shape3dBevelBottomHeightEmu` native leaf with the source coordinate

### Requirement: edit only the picture bottom-bevel height token

For a projected picture bottom-bevel height leaf, the presentation codec MUST
accept a source-bound edit only when the requested coordinate is canonical,
non-negative, within the signed 32-bit DrawingML range, and different from
the current coordinate. It MUST replace only the direct `a:bevelB/@h` lexical
value in the owning SlidePart. It MUST preserve the picture relationship,
crop, mask, effects, other direct 3-D attributes, bevel dimensions and
preset, and all non-target package parts.

#### Scenario: source-bound edit preserves picture topology

- **WHEN** a projected picture bottom-bevel height changes from `800` to
  `1600`
- **THEN** only the owning SlidePart changes, the edited `h` is `1600`, and
  the picture's other XML and package parts remain intact

### Requirement: keep unsupported picture 3-D owners opaque

The codec MUST keep a picture bottom-bevel owner source-bound when it has an
unknown or malformed height, duplicate `a:sp3d` or `a:bevelB` state, unknown
or extension attributes, extra children, or other unsupported 3-D topology.
It MUST NOT reconstruct or normalize such markup merely to expose or edit the
height leaf.

#### Scenario: reject an unsupported picture owner

- **WHEN** a picture's 3-D owner contains unsupported or ambiguous state
- **THEN** projection does not expose `shape3dBevelBottomHeightEmu` for that
  owner and a source-bound edit cannot target it
