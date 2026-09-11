## Why

Direct rich-text reflection now exposes positions, fade direction, and both scale axes plus horizontal skew, but a direct vertical skew value is still source-owned. A bounded `skewY` field closes the next native reflection scalar without opening arbitrary transform graphs.

## What Changes

- Add `run.style.reflection.skewY` and the `textReflectionSkewY` native leaf.
- Accept only a full-span direct run reflection with one canonical `ky` token and no other reflection transform or flag.
- Parse and project the signed 1/60000-degree value, and source-bound edit only `reflection/@ky` while preserving the rest of the package.
- Add the schema, capability registry, Skill/reference, coverage/backlog entry, OpenSpec artifacts, and a focused authored/source-bound regression.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-reflection-skew-y`: bounded authored and source-bound vertical skew for a direct rich-text reflection.

### Modified Capabilities

None.

## Impact

The PPJ schema/wire projection, NativeAOT reflection reader/projector/edit plan, presentation Skill references, generated capability matrix, and focused codec tests change. Unsupported effect graphs remain opaque and no PowerPoint host-rendering claim is added.
