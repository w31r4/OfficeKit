## Why

F-03/F-06 still preserve a direct vertical anchor when projected verticalAlignment is deleted. Explicit top/middle/bottom and absence must survive source edits independently of anchorCenter.

## What Changes

- Complete verticalAlignment deletion and restoration across existing text owners and table cells.
- Extend simple-style removal while retaining other-field and authority guards.
- Extend the shared enum experiment and synchronize field descriptions and preview limits.

## Capabilities

### New Capabilities
- `ppj-text-vertical-alignment-removal`: source-bound verticalAlignment presence lifecycle.

### Modified Capabilities

None.

## Impact

PPJ source compiler, bounded body admission, shared tests, schema, Help, registry, text reference and F-03/F-06 coverage. Reuses NoVerticalAnchor; existing middle-to-center normalization remains intact.
