# Proposal: expose one custom-geometry cubic second-control-point y coordinate

F-04 still keeps custom path command coordinates source-owned. Add one narrow
native leaf for the direct literal `a:cubicBezTo/a:pt[1]/@y` so an agent can
move the second control point of one cubic segment vertically without
rewriting the rest of the path graph.

## What Changes

- Expose each direct literal cubic second-control-point y coordinate as
  `customGeometryPathCubicControl2Y` in the owning path/command order.
- Allow a source-bound PPJ edit to replace only that point's `y` attribute in
  the owning SlidePart.
- Keep its x, the first control point, end point, path properties, other
  commands, sibling paths, and all other geometry source-owned.
- Add a focused source-bound/reprojection regression and update the F-04
  coverage record.

## Capabilities

### New Capabilities

- `ppj-source-bound-custom-geometry-path-cubic-control2-y`: a bounded native
  leaf for editing one direct cubic second-control-point y coordinate.

## Impact

- PPJ native-leaf projection, edit-plan validation/proof, and SlidePart XML
  token splicing.
- Capability registry, Skill reference, and PPJ gap/coverage documentation.
- Native codec focused tests; no protocol version change is required because
  native leaves already use the existing edit-plan wire shape.
