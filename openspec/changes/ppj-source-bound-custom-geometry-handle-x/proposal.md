## Why

F-04 already recognizes the standard custom-geometry adjustment-handle graph,
but the handle's direct position is still opaque to PPJ native editing. A
literal XY handle x coordinate is a small, independently addressable value;
exposing it closes one useful handle gap while preserving the controlled guide
identity and all handle topology.

## What Changes

- Add `customGeometryAdjustmentHandleXEmu` as a bounded PPJ native leaf for an
  ordered `a:ahXY` handle.
- Project only a direct, canonical, non-negative integer `a:ahXY/a:pos/@x`
  token inside the shape-local frame; formula-backed or malformed targets
  remain source-owned.
- Revalidate the handle index, XY handle kind, direct position structure,
  source token, and coordinate range before writing.
- Token-splice only the selected handle position's `@x` in the owning
  SlidePart; preserve `gdRefX/gdRefY`, bounds, y, polar siblings, paths,
  guides/formulas, and list topology.
- Add a focused authored/imported/source-bound/reprojection regression and
  update PPJ capability/reference coverage.

## Capabilities

### New Capabilities

- `ppj-source-bound-geometry-handle-x`: source-bound editing of one direct
  custom-geometry XY adjustment-handle x coordinate.

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
