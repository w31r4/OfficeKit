## Why

F-04 now exposes direct polar-handle radial bounds, but the polar angle range
is still opaque. A bounded `minAng` leaf lets an agent tune an existing angular
constraint while keeping the angle guide, maximum bound, radial state,
position, and custom-geometry topology source-faithful.

## What Changes

- Add `customGeometryAdjustmentHandlePolarMinAngle60000` as a bounded PPJ
  native leaf for an ordered `a:ahPolar` handle with a direct literal
  `@minAng`.
- Project only a canonical integer in the inclusive ±360-degree DrawingML
  range; reference-backed, missing, malformed, or non-polar targets remain
  source-owned.
- Revalidate the handle index/kind, direct handle attributes, source token, and
  angle bound before writing.
- Token-splice only the selected `a:ahPolar/@minAng` value in the owning
  SlidePart; preserve `gdRefR/gdRefAng`, radial bounds, `maxAng`, position,
  other handles, and custom-geometry topology.
- Add a focused authored/imported/source-bound/reprojection regression and
  update PPJ capability/reference coverage.

## Capabilities

### New Capabilities

- `ppj-source-bound-geometry-handle-polar-min-angle`: source-bound editing of
  one direct custom-geometry polar adjustment-handle minimum angle.

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

