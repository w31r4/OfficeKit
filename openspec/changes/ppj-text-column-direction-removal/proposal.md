## Why

F-03/F-06 still lack explicit deletion for text-body columnDirection. Source merging preserves rtlCol after the PPJ field is removed, preventing return to inherited/default direction.

## What Changes

- Preserve left-to-right, right-to-left and absence through source edits for existing editable text owners and table cells.
- Permit removing a style containing only columnDirection, upright and/or rotation.
- Reuse native NoColumnDirection and the shared lifecycle regression; document the public contract and preview limitation.

## Capabilities

### New Capabilities
- `ppj-text-column-direction-removal`: source-bound column direction presence lifecycle.

### Modified Capabilities

None.

## Impact

PPJ source compiler, bounded body-property admission, shared native tests, schema/Help/registry/references and gap documentation. Existing wire and capabilities remain sufficient.
