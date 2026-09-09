## Why

F-03/F-06 still preserve horizontal overflow after its projected field is deleted. Callers need explicit overflow/clip and absence to survive source edits and fresh projection.

## What Changes

- Complete horizontalOverflow removal/restoration on existing editable text owners and table cells.
- Extend simple-style deletion to this field while retaining other-field and authority guards.
- Extend the shared lifecycle experiment, public descriptions and preview diagnostic.

## Capabilities

### New Capabilities
- `ppj-text-horizontal-overflow-removal`: horizontalOverflow presence lifecycle under existing source edit authority.

### Modified Capabilities

None.

## Impact

PPJ source compiler, bounded body admission, native tests, schema, Help, registry, text reference and F-03/F-06 coverage. Existing NoHorizontalOverflowMode wire marker is sufficient.
