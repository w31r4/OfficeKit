# Proposal: expose one custom-geometry move-to x coordinate

F-04 still keeps custom path command coordinates source-owned. Add one narrow
native leaf for a direct literal `a:moveTo/a:pt/@x` so an agent can move the
start of one path without rewriting the rest of the path graph.

## What Changes

- Expose each direct literal move-to x coordinate as
  `customGeometryPathMoveToX` in the owning path/command order.
- Allow a source-bound PPJ edit to replace only that point's `x` attribute in
  the owning SlidePart.
- Keep the paired `y`, path properties, other commands, sibling paths, and all
  other geometry source-owned.
- Add a focused source-bound/reprojection regression and update the F-04
  coverage record.

## Capabilities

### New Capabilities

- `ppj-source-bound-custom-geometry-path-move-to-x`: a bounded native leaf for
  editing one direct move-to x coordinate.

## Impact

- PPJ native-leaf projection, edit-plan validation/proof, and SlidePart XML
  token splicing.
- Capability registry, Skill reference, and PPJ gap/coverage documentation.
- Native codec focused tests; no protocol version change is required because
  native leaves already use the existing edit-plan wire shape.
