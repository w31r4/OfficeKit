## Why

F-03/F-06 retain vertical overflow after its projected field is deleted. Callers need explicit overflow/ellipsis/clip and absence to survive source edits and fresh projection.

## What Changes

- Complete verticalOverflow removal and all three restored values across existing editable text owners and table cells.
- Extend simple-style deletion while retaining other-field and authority guards.
- Reuse the shared enum experiment, preserve horizontalOverflow independently, and update public descriptions and preview diagnostics.

## Capabilities

### New Capabilities
- `ppj-text-vertical-overflow-removal`: verticalOverflow source presence lifecycle.

### Modified Capabilities

None.

## Impact

PPJ source compiler, bounded body admission, shared tests, schema, Help, registry, text reference and F-03/F-06 coverage. The existing NoVerticalOverflowMode wire marker is sufficient.
