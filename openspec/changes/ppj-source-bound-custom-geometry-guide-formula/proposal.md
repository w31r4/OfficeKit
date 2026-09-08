# Proposal: expose one custom-geometry guide formula

F-04 still leaves calculated `a:gdLst` formulas source-owned. Add one narrow
native leaf for an existing canonical guide formula so an agent can change one
formula string while the rest of the custom-geometry graph remains intact.

## What Changes

- Expose each existing direct custom-geometry guide with a unique name and a
  canonical non-`val N` formula as `customGeometryGuideFormula`.
- Allow a source-bound PPJ edit to replace only that guide's `fmla` attribute
  value in the owning SlidePart.
- Revalidate the candidate formula against the existing bounded formula graph
  before applying it; do not edit names, references, handles, connection sites,
  text rectangles, paths, extensions, or graph topology.
- Add a focused authored/imported/source-bound/reprojection regression and
  update the F-04 coverage record.

## Capabilities

### New Capabilities

- `ppj-source-bound-geometry-guide-formula`: a bounded native leaf for editing
  one calculated custom-geometry guide formula.

## Impact

- PPJ native-leaf projection, edit-plan validation/proof, and SlidePart XML
  token splicing.
- Capability registry, Skill reference, and PPJ gap/coverage documentation.
- Native codec focused tests; no protocol version change is required because
  native leaves already use the existing edit-plan wire shape.
