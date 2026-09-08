## Why

F-04 already preserves and round-trips recognized custom-geometry paths, but an
explicit `a:path/@w` viewport width is still trapped inside the opaque path
topology. A bounded scalar leaf can expose this independent value without
pretending that path commands or formula-driven geometry are generally
editable.

## What Changes

- Add `customGeometryPathWidth` as a source-bound PPJ native leaf for each
  recognized custom-geometry path with a positive literal width.
- Bind the leaf's `NativeLeafIndex` to ordered path position and report the
  DrawingML path-coordinate value without converting it to EMU.
- Token-splice only the selected direct `a:path/@w`, preserving height,
  fill/stroke/extrusion attributes, commands, sibling paths, and package
  parts.
- Keep omitted, zero, negative, out-of-range, malformed, extension-bearing,
  and unsupported path topology source-owned.
- Add a focused authored/imported/source-bound/reprojection regression and
  update PPJ schema, capability, coverage, backlog, and generated reference
  evidence.

## Capabilities

### New Capabilities

- `ppj-source-bound-geometry-path-width`: source-bound editing of one explicit
  custom-geometry path viewport width.

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
