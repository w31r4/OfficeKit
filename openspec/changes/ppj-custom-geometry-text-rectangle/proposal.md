## Why

F-04 custom geometry already has a native text rectangle with literal/reference edges, but PPJ cannot author it and excludes shapes carrying it from its literal geometry projection. Expose that state without weakening the existing formula-graph boundary.

## What Changes

- Add geometry.textRectangle with left/top/right/bottom edges, in shape-local points or native built-in reference names.
- Preserve authored and fresh-projected literal/reference/mixed edges and omission.
- Extend setGeometry with independent text-rectangle add/change/removal on otherwise recognized literal custom shapes, preserving path and text state.
- Keep declared custom guide graphs opaque until their PPJ owner is implemented; retain explicit preview limitations.

## Capabilities

### New Capabilities
- `ppj-custom-geometry-text-rectangle`: Editable custom-shape text layout rectangle in PPJ.

### Modified Capabilities
None; the existing custom-path increment is unarchived.

## Impact

PPJ schema, authored compiler, source compiler/projector, existing native rectangle codec, focused tests and Help/registry/reference documentation. Native wire already supports all four edges; no protocol change.
