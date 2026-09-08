# Proposal: expose one custom-geometry arc sweep angle

F-04 already recognizes bounded `arcTo` path commands, but their scalar
attributes remain source-owned. Add one narrow native leaf for a direct
canonical `a:arcTo/@swAng` so an agent can change the signed sweep of one arc
within a single turn without rewriting the rest of the path graph.

## What Changes

- Expose each direct non-zero one-turn arc sweep angle as
  `customGeometryPathArcSweepAngle60000` in path/command order.
- Allow a source-bound PPJ edit to replace only that arc's `swAng` attribute in
  the owning SlidePart.
- Keep both radii, start angle, path properties, other commands, sibling paths,
  and formula/reference-backed values source-owned.
- Add a focused source-bound/reprojection regression and update the F-04
  coverage record.

## Capabilities

### New Capabilities

- `ppj-source-bound-custom-geometry-path-arc-sweep-angle`: a bounded native
  leaf for editing one direct custom-geometry arc sweep angle.

## Impact

- PPJ native-leaf projection, edit-plan validation/proof, and SlidePart XML
  token splicing.
- Capability registry, Skill reference, and PPJ gap/coverage documentation.
- Native codec focused tests; no protocol version change is required because
  native leaves already use the existing edit-plan wire shape.
