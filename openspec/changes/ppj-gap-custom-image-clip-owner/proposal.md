## Why

The bounded preset image clip profile now reaches the native picture mask, but
the same native owner also supports validated custom paths through `image.mask`.
Leaving custom paths rejected in `compositing.clipStack` keeps an existing PPJ
image-clip primitive split across two equivalent public spellings.

## What Changes

- Accept exactly one non-inverse custom-geometry clip for an image without a
  direct `mask` field when the existing custom-geometry codec accepts it.
- Lower the clip path graph to the existing native picture custom-mask owner.
- Project it back as canonical `image.mask.kind = "custom"` after the embedded
  PPJ snapshot is removed.
- Continue to reject multiple, inverse, preset-invalid, non-image, conflicting,
  or custom graphs outside the existing bounded path profile.
- Add focused XML/PPJ round-trip and diagnostic evidence, and update the PPJ
  reference, coverage, and K-04/F-05 backlog entries.

## Capabilities

### New Capabilities

- `ppj-custom-image-clip-owner`: A bounded single custom-path image clip that
  uses the existing native picture custom-mask owner.

### Modified Capabilities

- None.

## Impact

The PPJ semantic validator and authored presentation compiler will extend the
existing image clip predicate in `native/OfficeKit/src/OfficeKit.Codec/`.
`PresentationImage.CustomMaskPaths`, `PptxCustomGeometryCodec`, the wire
contract, and the PPJ schema remain unchanged. The native projector already
has the canonical custom `image.mask` representation. No host, provider, or
package changes are needed.
