## Why

F-04 already preserves and round-trips recognized custom-geometry paths, but an
explicit `a:path/@extrusionOk` choice is still trapped inside the opaque path
topology. The codec already preserves this optional boolean, so a narrow native
leaf can expose the choice without making the entire path graph editable.

## What Changes

- Add `customGeometryPathExtrusionAllowed` as a source-bound PPJ native leaf
  for each recognized custom-geometry path with an explicit `extrusionOk`
  token.
- Bind the leaf's `NativeLeafIndex` to ordered path position and map the
  explicit DrawingML boolean to a PPJ boolean.
- Token-splice only the selected direct `a:path/@extrusionOk`, preserving
  viewport extents, fill/stroke attributes, commands, sibling paths, and
  package parts.
- Keep omitted, malformed, extension-bearing, and unknown path topology
  source-owned.
- Add a focused authored/imported/source-bound/reprojection regression and
  update PPJ schema, capability, coverage, backlog, and generated reference
  evidence.

## Capabilities

### New Capabilities

- `ppj-source-bound-geometry-path-extrusion`: source-bound editing of one
  explicit custom-geometry path extrusion permission.

### Modified Capabilities

- None.

## Impact

- `PpjNativeLeafProjection` and `PptxEditPlanCodec` gain one bounded boolean
  native leaf; no protobuf or Office wire version change is required.
- The PPJ native-leaf schema and capability registry gain one enum/entry, and
  the generated Presentation Skill reference plus F-04 coverage evidence are
  refreshed.
- Native codec tests add a small PPTX round-trip fixture; no new dependency or
  host behavior is introduced.
