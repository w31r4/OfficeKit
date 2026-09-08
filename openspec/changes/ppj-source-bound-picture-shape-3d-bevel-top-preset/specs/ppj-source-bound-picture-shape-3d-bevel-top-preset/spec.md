# Picture 3-D top-bevel preset

## ADDED Requirements

### Requirement: project a bounded picture top-bevel preset

The presentation codec MUST project `p:pic/p:spPr/a:sp3d/a:bevelT/@prst`
through the existing `shape3dBevelTopPreset` native leaf when the picture
owner has exactly one direct `a:sp3d`, exactly one direct `a:bevelT`, only the
established direct 3-D scalar attributes, only `w`, `h`, and `prst` on
`bevelT`, and a recognized bevel-preset token.

#### Scenario: project a source-bound picture top-bevel preset

- **WHEN** a source PPTX contains a picture with a strict direct
  `a:sp3d/a:bevelT/@prst` owner
- **THEN** the projected picture contains exactly one
  `shape3dBevelTopPreset` native leaf with the source token

### Requirement: edit only the picture top-bevel preset token

For a projected picture top-bevel preset leaf, the presentation codec MUST
accept a source-bound edit only when the requested token is recognized and
different from the current token, and MUST replace only the direct
`a:bevelT/@prst` lexical value in the owning SlidePart. It MUST preserve the
picture relationship, crop, mask, effects, other direct 3-D attributes,
bevel dimensions, and all non-target package parts.

#### Scenario: source-bound edit preserves picture topology

- **WHEN** a projected picture top-bevel preset changes from `angle` to
  `softRound`
- **THEN** only the owning SlidePart changes, the edited `prst` is
  `softRound`, and the picture's other XML and package parts remain intact

### Requirement: keep unsupported picture 3-D owners opaque

The codec MUST keep a picture top-bevel owner source-bound when it has an
unknown or malformed preset, duplicate `a:sp3d` or `a:bevelT` state, unknown
or extension attributes, extra children, or other unsupported 3-D topology.
It MUST NOT reconstruct or normalize such markup merely to expose or edit
the preset leaf.

#### Scenario: reject an unsupported picture owner

- **WHEN** a picture's 3-D owner contains unsupported or ambiguous state
- **THEN** projection does not expose `shape3dBevelTopPreset` for that owner
  and a source-bound edit cannot target it
