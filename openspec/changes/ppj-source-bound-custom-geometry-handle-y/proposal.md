## Why

F-04 now has a narrow source-bound leaf for a literal XY adjustment-handle x
position, but the matching y position remains opaque. Adding the y leaf closes
the other direct coordinate without exposing guide identity, range mutation,
formula evaluation, or handle topology changes.

## What Changes

- Add `customGeometryAdjustmentHandleYEmu` as a bounded PPJ native leaf for an
  ordered `a:ahXY` handle.
- Project only a direct, canonical, non-negative integer
  `a:ahXY/a:pos/@y` token inside the shape-local frame; formula-backed or
  malformed targets remain source-owned.
- Revalidate the handle index, XY handle kind, direct position structure,
  source token, and coordinate range before writing.
- Token-splice only the selected handle position's `@y` in the owning
  SlidePart; preserve `gdRefX/gdRefY`, bounds, x, polar siblings, paths,
  guides/formulas, and list topology.
- Add a focused authored/imported/source-bound/reprojection regression and
  update PPJ capability/reference coverage.

## Capabilities

### New Capabilities

- `ppj-source-bound-geometry-handle-y`: source-bound editing of one direct
  custom-geometry XY adjustment-handle y coordinate.

### Modified Capabilities

- None.

## Impact

- `PpjNativeLeafProjection` and `PptxEditPlanCodec` gain one generic native
  scalar leaf; no protobuf or Office wire version change is required.
- The PPJ native-leaf schema and capability registry gain one enum/entry, and
  the generated Presentation Skill reference plus F-04 coverage evidence are
  refreshed.
- Native codec tests add a small PPTX round-trip fixture; no new dependency or
  host behavior is introduced.
