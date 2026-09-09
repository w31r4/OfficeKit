## Why

F-03/F-06 retain a text direction after its projected verticalText field is removed. Completing this lifecycle lets source edits return to inherited/default direction while keeping an explicit horizontal setting distinct.

## What Changes

- Support removal and restoration of horizontal, vertical and vertical270 for existing editable text owners and table cells.
- Extend removable simple styles to verticalText alongside upright, rotation and columnDirection.
- Reuse the native deletion marker and focused lifecycle experiment; synchronize public descriptions and preview diagnostics.

## Capabilities

### New Capabilities
- `ppj-text-vertical-direction-removal`: verticalText source-bound presence lifecycle.

### Modified Capabilities

None.

## Impact

PPJ source compiler, bounded body-property admission, shared lifecycle tests, schema, Help, registry, references and F-03/F-06 documentation. Existing protocol and edit authority are sufficient.
