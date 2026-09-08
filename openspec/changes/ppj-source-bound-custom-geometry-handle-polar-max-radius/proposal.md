## Why

F-04 now exposes a direct polar-handle minimum-radius bound, but the paired
`maxR` bound remains opaque. A bounded `maxR` leaf lets an agent tune the
existing radial constraint while keeping the radial guide, minimum bound,
angle state, position, and custom-geometry topology source-faithful.

## What Changes

- Add `customGeometryAdjustmentHandlePolarMaxRadiusEmu` as a bounded PPJ
  native leaf for an ordered `a:ahPolar` handle with a direct literal `@maxR`.
- Project only a canonical non-negative integer radial bound within the
  DrawingML integer range; reference-backed, missing, malformed, or non-polar
  targets remain source-owned.
- Revalidate the handle index/kind, direct handle attributes, source token, and
  radial bound before writing.
- Token-splice only the selected `a:ahPolar/@maxR` value in the owning
  SlidePart; preserve `gdRefR/gdRefAng`, `minR`, angle bounds, position, other
  handles, and custom-geometry topology.
- Add a focused authored/imported/source-bound/reprojection regression and
  update PPJ capability/reference coverage.

## Capabilities

### New Capabilities

- `ppj-source-bound-geometry-handle-polar-max-radius`: source-bound editing of
  one direct custom-geometry polar adjustment-handle maximum radius.

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

