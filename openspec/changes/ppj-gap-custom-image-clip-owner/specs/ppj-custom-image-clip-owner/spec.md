## Purpose

This capability lets an authored image use one validated custom path as its
compositing clip while retaining the existing editable native picture mask.

## ADDED Requirements

### Requirement: Authored PPJ SHALL lower one bounded custom image clip

An authored image with exactly one non-inverse `compositing.clipStack` entry
whose geometry is an existing bounded custom path SHALL be accepted when it
has no direct `mask` field. The path SHALL become the image's native custom
mask owner and SHALL remain editable as path data.

#### Scenario: Custom clip authors as a picture custom geometry

- **WHEN** an image declares one non-inverse custom clip with a valid view box
  and path command graph
- **THEN** authored compilation succeeds and the picture shape contains the
  corresponding custom geometry without rasterizing the image

### Requirement: Custom image clip projection SHALL use the canonical mask

When a native picture contains the custom geometry produced by the bounded
clip profile, PPJ projection SHALL expose it as `image.mask.kind = "custom"`
with its path commands. It SHALL not invent a multi-entry clip stack.

#### Scenario: Snapshot-free projection recovers custom paths

- **WHEN** the authored PPTX is projected after its embedded PPJ snapshot is
  removed
- **THEN** the image contains the same bounded custom mask path graph and no
  clip-unsupported diagnostic

### Requirement: Unsupported custom clip topology SHALL fail closed

The implementation SHALL reject inverse or multiple custom clips, custom
geometry outside the existing native profile, non-image owners, and a direct
image `mask` combined with a clip. It SHALL report
`ppj.compositing.clipUnsupported` and SHALL not rasterize or silently discard
the rejected path graph.

#### Scenario: Custom clip combinations outside one owner are refused

- **WHEN** a program supplies an inverse custom clip, two clip entries, an
  unsupported custom command graph, or both `mask` and `clipStack`
- **THEN** validation fails with a clip-unsupported diagnostic and no lossy
  authored replacement is produced
