## Why

F-04 can preserve and round-trip a custom-geometry text rectangle, and the
left edge now has a bounded source-bound leaf, but the resolved top edge is
still opaque. A matching top-edge leaf lets an agent adjust the vertical text
inset while keeping the other edges, guide representation, and geometry
topology source-faithful.

## What Changes

- Add `customGeometryTextRectangleTopEmu` as a bounded PPJ native leaf for a
  recognized custom-geometry text rectangle whose top edge is literal.
- Accept either a canonical direct numeric `a:rect/@t` token or the exact
  OfficeKit literal guide profile (`t="officeKitTextTop"` with its canonical
  scaled `a:gd` formula); reference-backed and malformed edges remain
  source-owned.
- Require all four rectangle edges to be canonical in-frame literals and the
  resolved rectangle to remain ordered before issuing the leaf.
- Token-splice only the top coordinate: direct sources change `a:rect/@t`,
  while the private profile changes only the matching guide formula and keeps
  the `rect` reference intact.
- Add a focused authored/imported/source-bound/reprojection regression and
  update PPJ capability/reference coverage.

## Capabilities

### New Capabilities

- `ppj-source-bound-geometry-text-rectangle-top`: source-bound editing of one
  direct custom-geometry text-rectangle top coordinate.

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
