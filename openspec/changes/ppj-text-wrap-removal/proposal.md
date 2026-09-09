## Why

F-03/F-06 source edits still retain native wrapping after the projected wrap field is removed. Explicit none and absence must remain distinct so callers can return to inherited/default layout.

## What Changes

- Complete square/none/removal/restoration for supported text owners and table cells.
- Allow removing simple styles containing wrap and the existing removable body fields.
- Extend the shared lifecycle experiment and public field descriptions; retain explicit preview layout limitations.

## Capabilities

### New Capabilities
- `ppj-text-wrap-removal`: direct text wrap presence lifecycle under existing source edit authority.

### Modified Capabilities

None.

## Impact

PPJ source compiler, bounded body-property admission, lifecycle tests, schema/Help/registry/references and F-03/F-06 coverage. Reuses the existing NoWrap wire marker; no protocol revision.
