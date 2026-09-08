## Why

F-04 can preserve and round-trip a custom-geometry text rectangle, but its
resolved edges remain opaque to source-bound editing. A bounded left-edge leaf
lets an agent widen or narrow an existing text rectangle while keeping the
other edges, private guide profile, path graph, and custom-geometry topology
source-faithful.

## What Changes

- Add `customGeometryTextRectangleLeftEmu` as a bounded PPJ native leaf for a
  recognized custom-geometry text rectangle whose left edge is literal.
- Accept either a canonical direct numeric `a:rect/@l` token or the exact
  OfficeKit literal guide profile (`l="officeKitTextLeft"` with its canonical
  scaled `a:gd` formula); reference-backed and malformed edges remain
  source-owned.
- Require all four rectangle edges to be canonical in-frame literals and the
  resolved rectangle to remain ordered before issuing the leaf.
- Token-splice only the left coordinate: direct sources change `a:rect/@l`,
  while the private profile changes only the corresponding guide formula and
  keeps the `rect` reference intact.
- Add a focused authored/imported/source-bound/reprojection regression and
  update PPJ capability/reference coverage.

## Capabilities

### New Capabilities

- `ppj-source-bound-geometry-text-rectangle-left`: source-bound editing of one
  direct custom-geometry text-rectangle left coordinate.

### Modified Capabilities

- None.

## Impact

- `PpjNativeLeafProjection` and `PptxEditPlanCodec` gain one bounded numeric
  native leaf; no protobuf or Office wire version change is required.
- The PPJ native-leaf schema and capability registry gain one enum/entry, and
  the generated Presentation Skill reference plus F-04 coverage evidence are
  refreshed.
- Native codec tests add a small PPTX round-trip fixture; no new dependency or
  host behavior is introduced.
