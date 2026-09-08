## Purpose

This capability gives PPJ a small, editable compositing-clip profile by using
the single native mask owner already available on a PowerPoint picture.

## ADDED Requirements

### Requirement: Authored PPJ SHALL lower a bounded single image clip

An authored image element with exactly one `compositing.clipStack` item SHALL
be accepted when the item is non-inverse preset geometry, its literal
adjustments satisfy the canonical preset profile, and the image has no direct
`mask` field. The clip SHALL lower to the image's native preset mask owner.

#### Scenario: Single preset clip authors as a picture mask

- **WHEN** an image declares `compositing.clipStack` with one clip whose
  `inverse` is omitted or `false`, whose geometry is a supported preset such as
  `roundRect`, and whose adjustments are valid for that preset
- **THEN** authored compilation succeeds and the picture shape contains the
  corresponding preset geometry and ordered adjustment values

#### Scenario: Clip and normal opacity share the image owner safely

- **WHEN** the same image declares the bounded single clip and
  `compositing.opacity`
- **THEN** compilation writes both the picture mask and the image opacity,
  without flattening the image or changing its asset relationship

### Requirement: Native image-mask projection SHALL be the canonical recovery

When a native picture contains the mask produced by the bounded clip profile,
PPJ projection SHALL expose it as `image.mask` with its preset and complete
adjustment list. The projection SHALL not invent a multi-entry clip stack from
one native mask owner.

#### Scenario: Snapshot-free projection recovers the canonical mask

- **WHEN** the authored PPTX is projected after its embedded PPJ snapshot is
  removed
- **THEN** the image contains `mask.kind = "preset"` with the same preset and
  adjustment values, and no unsupported clip-stack diagnostic is emitted

### Requirement: Unsupported clip topology SHALL fail closed

The implementation SHALL reject a non-empty clip stack unless it matches the
bounded single-image profile. It SHALL report `ppj.compositing.clipUnsupported`
for multiple clips, inverse clips, custom geometry, non-image owners,
unsupported presets or adjustments, and a direct image `mask` combined with a
clip. It SHALL not silently flatten or discard those semantics.

#### Scenario: Broader clip forms are refused

- **WHEN** a PPJ element uses two clip entries, an inverse entry, a custom
  geometry entry, a non-image element, or an image that already has `mask`
- **THEN** semantic validation fails with a clip-unsupported diagnostic and
  authored compilation does not produce a lossy replacement
