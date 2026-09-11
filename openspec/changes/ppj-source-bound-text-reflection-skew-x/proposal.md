## Why

Direct rich-text reflection now round-trips positions, fade direction, and both scale axes, but a direct horizontal skew value is still source-owned. A bounded `skewX` field closes the next native reflection scalar without opening arbitrary transform graphs.

## What Changes

- Add `run.style.reflection.skewX` and the `textReflectionSkewX` native leaf.
- Accept only a full-span direct run reflection with one canonical `kx` token and no other reflection transform or flag.
- Parse and project the signed 1/60000-degree value, and source-bound edit only `reflection/@kx` while preserving the rest of the package.
- Add schema, capability registry, generated references, coverage/backlog wording, OpenSpec artifacts, and a focused regression.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-reflection-skew-x`: bounded authored and source-bound horizontal skew for a direct rich-text reflection.

### Modified Capabilities

None.

## Impact

The PPJ schema/projection, NativeAOT reflection reader and edit plan, presentation Skill references, generated capability matrix, and focused codec tests change. Unsupported effect graphs remain opaque and no PowerPoint host-rendering claim is added.
