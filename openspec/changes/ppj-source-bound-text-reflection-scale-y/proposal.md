## Why

Direct rich-text reflection already exposes bounded position, fade direction, and horizontal scale leaves, but the vertical scale value remains source-owned. That leaves one native reflection scalar without a complete PPJ field and prevents a minimal single-attribute edit from round-tripping.

## What Changes

- Add `run.style.reflection.scaleY` and the `textReflectionScaleY` native leaf.
- Accept only a full-span direct run reflection with one canonical `sy` token and no other reflection transform or flag.
- Parse and project the signed native 1/100000 ratio, and source-bound edit only `reflection/@sy` while preserving the rest of the package.
- Add the schema, capability registry, Skill/reference, coverage/backlog entry, OpenSpec artifacts, and a focused authored/source-bound regression.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-reflection-scale-y`: bounded authored and source-bound vertical scale for a direct rich-text reflection.

### Modified Capabilities

None.

## Impact

The PPJ schema/wire projection, NativeAOT text reflection reader/projector/edit plan, presentation Skill references, generated capability matrix, and focused codec tests change. Unsupported reflection graphs remain opaque and no PowerPoint host-rendering claim is added.
