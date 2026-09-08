## Why

F-04 now exposes direct polar-handle range bounds, but a polar handle's
shape-local position remains opaque. A bounded position-x leaf lets an agent
adjust the existing handle anchor while keeping its guide references, y
coordinate, range graph, and custom-geometry topology source-faithful.

## What Changes

- Add `customGeometryAdjustmentHandlePolarXEmu` as a bounded PPJ native leaf
  for an ordered `a:ahPolar/a:pos` with a direct literal `@x`.
- Project only a canonical non-negative integer within the shape-local width;
  reference-backed, missing, malformed, or non-polar targets remain
  source-owned.
- Revalidate the handle index/kind, direct handle attributes and position,
  source token, and shape-local bound before writing.
- Token-splice only the selected `a:pos/@x` value in the owning SlidePart;
  preserve `gdRefR/gdRefAng`, radial and angular bounds, y, other handles, and
  custom-geometry topology.
- Add a focused authored/imported/source-bound/reprojection regression and
  update PPJ capability/reference coverage.

## Capabilities

### New Capabilities

- `ppj-source-bound-geometry-handle-polar-x`: source-bound editing of one
  direct custom-geometry polar adjustment-handle x coordinate.

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
