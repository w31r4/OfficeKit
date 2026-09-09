## Why

F-04 native geometry has an ordered formula evaluator, but PPJ drops shapes with custom guides into a source-owned profile. This prevents an author from defining a named formula used by the newly exposed text rectangle.

## What Changes

- Add geometry.guides as an ordered list of name/formula objects for custom shapes.
- Preserve direct guide formulas and references through authored export and fresh projection, including textRectangle references to prior guides.
- Extend setGeometry to add, edit and remove the guide list, with graph validation and unchanged independent shape state.
- Keep unsupported adjustment/handle/site graphs and reference-backed paths source-owned until their PPJ fields exist; masks/clips reject shape-only guides.

## Capabilities

### New Capabilities
- `ppj-custom-geometry-guides`: Ordered named geometry formulas with PPJ source-edit lifecycle.

### Modified Capabilities
None; custom-path and text-rectangle increments remain unarchived.

## Impact

PPJ schema/compiler/projector/capabilities, existing native formula graph validation, one regression, Help/registry/reference and explicit preview diagnostics. Existing wire guide messages suffice.
