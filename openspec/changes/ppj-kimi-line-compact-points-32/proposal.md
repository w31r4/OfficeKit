## Why

The bounded Kimi `line.points` authoring path already lowers up to 128 points,
but snapshot-free projection stops preserving the same compact spelling after
24 points. This leaves a directly supported authored shape needlessly exposed
as a typed path after import.

## What Changes

- Extend the compact smooth-point reverse projection from 24 to 32 points.
- Use a numerically safer bounded least-squares solve for the higher-degree
  reverse projection, without loosening the existing proof tolerance.
- Keep the projection bounded and formula-checked; paths beyond 32 points or
  with quantization that cannot be recovered remain typed paths.
- Add one 25-point authored/imported/reprojection regression and update the
  K-01 evidence boundary.

## Capabilities

### New Capabilities

None. This is a bounded extension of the existing Kimi line profile.

### Modified Capabilities

None. The PPJ schema and wire contract already expose `line.points`.

## Impact

- `native/OfficeKit/src/OfficeKit.Codec/PpjLinePathCodec.cs`
- `native/OfficeKit/tests/OfficeKit.Codec.Tests/PptxCodecTests.cs`
- K-01 backlog evidence

No protocol, package, relationship, or full DrawingML geometry behavior
changes.
