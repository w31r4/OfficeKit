## Why

F-04 now has narrow source-bound leaves for a literal connection-site angle
and x coordinate, but a direct site position's y coordinate is still opaque.
Adding the matching y leaf closes the second half of the position without
making PPJ responsible for formula evaluation, connector rerouting, or
connection-site topology.

## What Changes

- Add `customGeometryConnectionSiteYEmu` as a bounded PPJ native leaf for one
  ordered custom-geometry connection site.
- Project only a direct, canonical, non-negative integer `a:pos/@y` token whose
  value is inside the shape-local frame; formula-backed or malformed targets
  remain source-owned.
- Revalidate the site index, source token, coordinate range, and direct XML
  structure before writing.
- Token-splice only the selected `a:pos/@y` value in the owning SlidePart;
  preserve x, angle, guides/formulas, handles, paths, site order, and package
  topology.
- Add a focused authored/imported/source-bound/reprojection regression and
  update PPJ capability/reference coverage.

## Capabilities

### New Capabilities

- `ppj-source-bound-geometry-connection-y`: source-bound editing of one direct
  custom-geometry connection-site y coordinate.

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
