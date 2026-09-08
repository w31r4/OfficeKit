## Why

F-04 now exposes direct polar-handle range bounds and both the polar anchor's
x coordinate, but its shape-local y coordinate remains opaque. A bounded
position-y leaf lets an agent adjust the existing handle anchor while keeping
its guide references, x coordinate, range graph, and custom-geometry topology
source-faithful.

## What Changes

- Add `customGeometryAdjustmentHandlePolarYEmu` as a bounded PPJ native leaf
  for an ordered `a:ahPolar/a:pos` with a direct literal `@y`.
- Project only a canonical non-negative integer within the shape-local height;
  reference-backed, missing, malformed, or non-polar targets remain
  source-owned.
- Revalidate the handle index/kind, direct handle attributes and position,
  source token, and shape-local bound before writing.
- Token-splice only the selected `a:pos/@y` value in the owning SlidePart;
  preserve `gdRefR/gdRefAng`, radial and angular bounds, x, other handles, and
  custom-geometry topology.
- Add a focused authored/imported/source-bound/reprojection regression and
  update PPJ capability/reference coverage.

## Capabilities

### New Capabilities

- `ppj-source-bound-geometry-handle-polar-y`: source-bound editing of one
  direct custom-geometry polar adjustment-handle y coordinate.

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
