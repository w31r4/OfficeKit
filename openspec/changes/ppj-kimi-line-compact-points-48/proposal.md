## Why

The existing Kimi `line.points` reverse projection now covers the first
post-24 bounded range, but authored smooth point lists can still become typed
paths above that boundary even when the native cubic chain is exact. A small
next range reduces this gap while preserving the explicit numeric safety
boundary.

## What Changes

- Extend compact smooth-point reverse projection from 32 to 48 points.
- Keep Householder QR, formula checks, quantization tolerance, and fail-closed
  behavior for unstable or arbitrary paths.
- Add one 33-point authored/imported/reprojection regression and update K-01
  evidence.

## Capabilities

### New Capabilities

None. This extends the existing Kimi line profile.

### Modified Capabilities

None. The PPJ schema and wire contract already expose `line.points`.

## Impact

- `native/OfficeKit/src/OfficeKit.Codec/PpjLinePathCodec.cs`
- `native/OfficeKit/tests/OfficeKit.Codec.Tests/PptxCodecTests.cs`
- K-01 backlog evidence

No protocol, package, relationship, or full DrawingML geometry behavior
changes.
