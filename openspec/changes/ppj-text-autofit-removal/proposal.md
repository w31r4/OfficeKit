## Why

F-03/F-06 source edits retain AutoFit choices and normalAutoFit percentages after PPJ deletes them. Explicit none and absence need distinct behavior.

## What Changes

- Complete autoFit removal/restoration and normalAutoFit independent percentage removal for supported text owners and tables.
- Retain shrink-text dependency validation and noncanonical source boundaries.
- Reuse lifecycle experiments; publish direct-field semantics and preview limitations.

## Capabilities

### New Capabilities
- `ppj-text-autofit-removal`: AutoFit choice and percentage presence lifecycle.

### Modified Capabilities

None.

## Impact

PPJ source compiler, bounded body admission, shared lifecycle tests and public field descriptions. Existing NoAutoFitMode/NoFontScale/NoLineSpacingReduction markers suffice; no wire change.
