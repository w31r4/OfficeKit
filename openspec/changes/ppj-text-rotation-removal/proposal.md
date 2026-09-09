## Why

F-03/F-06 body-style rotation can be authored and changed but source omission preserves the old native rot attribute. Complete this field's lifecycle, following the upright deletion contract.

## What Changes

- Remove direct bodyPr rot when a previously projected text-body rotation disappears, retaining explicit zero and signed degree values.
- Support text, shape, editable master/layout placeholder and structured table styles through existing capabilities.
- Permit removing a style owner containing only removable upright/rotation fields; retain other whole-style removal guards.
- Preserve other body state and package members, including table compact-text normalization and restoration after the last property is removed.

## Capabilities

### New Capabilities
- `ppj-text-rotation-removal`: Text-body rotation presence lifecycle in source PPJ.

### Modified Capabilities
None.

## Impact

Source body merge/table builder, native bounded NoRotation admission, focused lifecycle tests, schema/Help/registry/text reference and gap coverage. Existing wire deletion operation suffices; no host-layout acceptance.
