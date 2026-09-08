# Picture 3-D scene camera preset

## ADDED Requirements

### Requirement: project a bounded picture 3-D scene camera preset

The presentation codec MUST project
`p:pic/p:spPr/a:scene3d/a:camera/@prst` through the existing
`shape3dSceneCameraPreset` native leaf when the picture owner has exactly one
direct `a:scene3d` with exactly two direct children in camera/light-rig order,
a recognized camera preset with only the canonical optional camera attributes,
and a recognized light-rig pair.

#### Scenario: project a source-bound picture camera preset

- **WHEN** a source PPTX contains a picture with a strict direct
  `a:scene3d/a:camera/@prst` owner
- **THEN** the projected picture contains exactly one
  `shape3dSceneCameraPreset` native leaf with the canonical camera-preset
  value

### Requirement: edit only the picture camera preset token

For a projected picture camera-preset leaf, the presentation codec MUST
accept a source-bound edit only when the requested value is a recognized,
changed camera-preset token, and MUST replace only the direct
`a:scene3d/a:camera/@prst` lexical value in the owning SlidePart. It MUST
preserve the picture relationship, crop, mask, effects, other 3-D attributes,
the light rig, and all non-target package parts.

#### Scenario: source-bound edit preserves picture topology

- **WHEN** a projected picture camera preset changes from `perspectiveFront`
  to `orthographicFront`
- **THEN** only the owning SlidePart changes, the edited `prst` is
  `orthographicFront`, and the picture's other XML and package parts remain
  intact

### Requirement: keep unsupported picture scene owners opaque

The codec MUST keep a picture scene camera-preset owner source-bound when it
has an unknown or malformed preset, duplicate scene/camera/light-rig state,
unknown or extension attributes, extra children, partial transforms, or other
unsupported 3-D topology. It MUST NOT reconstruct or normalize such markup
merely to expose or edit the camera-preset leaf.

#### Scenario: reject an unsupported picture scene owner

- **WHEN** a picture's scene camera owner contains unsupported or ambiguous
  state
- **THEN** projection does not expose `shape3dSceneCameraPreset` for that
  owner and a source-bound edit cannot target it
