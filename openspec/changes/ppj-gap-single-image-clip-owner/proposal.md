## Why

The PPJ v1 grammar already exposes `compositing.clipStack`, but every non-empty
stack is currently rejected even when it describes the one clip that a native
PowerPoint picture already owns as its mask geometry. This leaves a common
single-image crop/clip semantic unavailable to authored PPJ and makes the
existing native mask contract harder to reach from the public language.

## What Changes

- Accept one non-inverse preset-geometry entry in `compositing.clipStack` for an
  image that does not also declare the element-level `mask` field.
- Lower that bounded clip to the existing native picture mask owner, including
  the preset's ordered literal adjustment values.
- Project the native owner back as canonical `image.mask`; do not claim that
  the original multi-entry compositing stack is recoverable from a native
  picture mask.
- Keep multi-entry, inverse, custom-geometry, non-image, conflicting-mask, and
  otherwise unsupported clip forms explicitly fail-closed.
- Add a minimal authored PPTX/XML/PPJ round-trip test and update the PPJ
  reference, coverage entry, and K-04 backlog evidence.

## Capabilities

### New Capabilities

- `ppj-single-image-clip-owner`: A bounded single-image compositing clip that
  uses the existing native picture mask owner.

### Modified Capabilities

- None.

## Impact

The PPJ semantic validator and authored presentation compiler will gain a
small lowering path in `native/OfficeKit/src/OfficeKit.Codec/`. The existing
`PresentationImage` mask fields, DrawingML picture codec, wire contract, and
schema remain unchanged. The projector's existing image-mask projection
becomes the canonical output for this authored clip. Tests and PPJ/K-04
documentation receive focused evidence; no provider, host, or package change
is required.
