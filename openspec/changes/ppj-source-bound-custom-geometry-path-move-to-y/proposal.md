# Proposal: expose one custom-geometry move-to y coordinate

F-04 still keeps custom path command coordinates source-owned. Add one narrow
native leaf for a direct literal `a:moveTo/a:pt/@y` so an agent can move the
start of one path vertically without rewriting the rest of the path graph.

## What Changes

- Expose each direct literal move-to y coordinate as
  `customGeometryPathMoveToY` in the owning path/command order.
- Allow a source-bound PPJ edit to replace only that point's `y` attribute in
  the owning SlidePart.
- Keep the paired `x`, path properties, other commands, sibling paths, and all
  other geometry source-owned.
- Add a focused source-bound/reprojection regression and update the F-04
  coverage record.

## Capabilities

### New Capabilities

- `ppj-source-bound-custom-geometry-path-move-to-y`: a bounded native leaf for
  editing one direct move-to y coordinate.

## Impact

- PPJ native-leaf projection, edit-plan validation/proof, and SlidePart XML
  token splicing.
- Capability registry, Skill reference, and PPJ gap/coverage documentation.
- Native codec focused tests; no protocol version change is required because
  native leaves already use the existing edit-plan wire shape.
