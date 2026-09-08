## Why

F-04 now exposes the direct x and y positions of an XY adjustment handle and
the x-range bounds, but the paired `minY` bound remains opaque. A bounded
`minY` leaf lets an agent tune the existing handle constraint while keeping
the controlled guide identity, current geometry, and the rest of the range
graph source-faithful.

## What Changes

- Add `customGeometryAdjustmentHandleMinYEmu` as a bounded PPJ native leaf for
  an ordered `a:ahXY` handle with a direct literal `@minY`.
- Project only a canonical non-negative integer bound inside the shape-local
  frame; reference-backed, missing, malformed, or non-XY targets remain
  source-owned.
- Revalidate the handle index/kind, direct handle attributes, source token, and
  shape-frame range before writing.
- Token-splice only the selected `a:ahXY/@minY` value in the owning SlidePart;
  preserve `gdRefX/gdRefY`, x bounds, `maxY`, position, other handles, and
  custom-geometry topology.
- Add a focused authored/imported/source-bound/reprojection regression and
  update PPJ capability/reference coverage.

## Capabilities

### New Capabilities

- `ppj-source-bound-geometry-handle-min-y`: source-bound editing of one direct
  custom-geometry XY adjustment-handle minimum-y bound.

### Modified Capabilities

- None.

## Impact

- `PpjNativeLeafProjection` and `PptxEditPlanCodec` gain one generic numeric
  native leaf; no protobuf or Office wire version change is required.
- The PPJ native-leaf schema and capability registry gain one enum/entry, and
  the generated Presentation Skill reference plus F-04 coverage evidence are
  refreshed.
- Native codec tests add a small PPTX round-trip fixture; no new dependency or
  host behavior is introduced.

