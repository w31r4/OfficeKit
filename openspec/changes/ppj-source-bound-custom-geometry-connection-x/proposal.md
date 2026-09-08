## Why

F-04 already preserves recognized custom-geometry connection sites, but PPJ
cannot independently address a literal site position. That leaves a small,
source-bound coordinate gap even when the site list and all other geometry are
safe to retain. Exposing one direct `a:pos/@x` value gives agents a useful
coordinate-level edit without pretending to support connection-site topology
or formula evaluation.

## What Changes

- Add `customGeometryConnectionSiteXEmu` as a bounded PPJ native leaf for one
  ordered custom-geometry connection site.
- Project only a direct, canonical, non-negative integer `a:pos/@x` token whose
  value is inside the shape-local frame; formula-backed or malformed targets
  remain source-owned.
- Revalidate the site index, source token, coordinate range, and direct XML
  structure before writing.
- Token-splice only the selected `a:pos/@x` attribute in the owning SlidePart;
  preserve y, angle, guides/formulas, handles, paths, site order, and package
  topology.
- Add a focused authored/imported/source-bound/reprojection regression and
  update PPJ capability/reference coverage.

## Capabilities

### New Capabilities

- `ppj-source-bound-geometry-connection-x`: source-bound editing of one direct
  custom-geometry connection-site x coordinate.

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
