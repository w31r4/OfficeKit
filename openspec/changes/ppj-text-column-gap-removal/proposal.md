## Why

F-03/F-06 retain column spacing when projected columnGap is deleted. Explicit zero spacing must remain distinct from removal of the direct override.

## What Changes

- Complete columnGap deletion and numeric restoration on existing editable text owners and table cells.
- Extend simple-style deletion to columnGap while retaining other-field and authority guards.
- Verify zero, fractional and boundary values in the shared lifecycle fixture and synchronize public descriptions and preview limits.

## Capabilities

### New Capabilities
- `ppj-text-column-gap-removal`: source-bound columnGap presence lifecycle.

### Modified Capabilities

None.

## Impact

PPJ source compiler, bounded body admission, shared native fixture, schema, Help, registry, text reference and F-03/F-06 coverage. Reuses NoColumnSpacing and existing points-to-EMU conversion.
