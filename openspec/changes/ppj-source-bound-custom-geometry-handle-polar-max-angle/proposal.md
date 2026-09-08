## Why

F-04 now exposes direct polar-handle radial bounds and the minimum angle, but
the opposite angle bound is still opaque. A bounded `maxAng` leaf lets an agent
tune an existing angular constraint while keeping the angle guide, minimum
bound, radial state, position, and custom-geometry topology source-faithful.

## What Changes

- Add `customGeometryAdjustmentHandlePolarMaxAngle60000` as a bounded PPJ
  native leaf for an ordered `a:ahPolar` handle with a direct literal
  `@maxAng`.
- Project only a canonical integer in the inclusive ±360-degree DrawingML
  range; reference-backed, missing, malformed, or non-polar targets remain
  source-owned.
- Revalidate the handle index/kind, direct handle attributes, source token, and
  angle bound before writing.
- Token-splice only the selected `a:ahPolar/@maxAng` value in the owning
  SlidePart; preserve `gdRefR/gdRefAng`, `minAng`, radial bounds, position,
  other handles, and custom-geometry topology.
- Add a focused authored/imported/source-bound/reprojection regression and
  update PPJ capability/reference coverage.

## Capabilities

### New Capabilities

- `ppj-source-bound-geometry-handle-polar-max-angle`: source-bound editing of
  one direct custom-geometry polar adjustment-handle maximum angle.

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
