# Proposal: expose one custom-geometry arc height radius

F-04 already recognizes bounded `arcTo` path commands, but their scalar
attributes remain source-owned. Add one narrow native leaf for direct literal
`a:arcTo/@hR` so an agent can change the vertical radius of one arc without
rewriting the rest of the path graph.

## What Changes

- Expose each direct positive literal arc height radius as
  `customGeometryPathArcHeightRadius` in path/command order.
- Allow a source-bound PPJ edit to replace only that arc's `hR` attribute in
  the owning SlidePart.
- Keep `wR`, start/sweep angles, path properties, other commands, sibling
  paths, and formula/reference-backed values source-owned.
- Add a focused source-bound/reprojection regression and update the F-04
  coverage record.

## Capabilities

### New Capabilities

- `ppj-source-bound-custom-geometry-path-arc-height-radius`: a bounded native
  leaf for editing one direct custom-geometry arc height radius.

## Impact

- PPJ native-leaf projection, edit-plan validation/proof, and SlidePart XML
  token splicing.
- Capability registry, Skill reference, and PPJ gap/coverage documentation.
- Native codec focused tests; no protocol version change is required because
  native leaves already use the existing edit-plan wire shape.
