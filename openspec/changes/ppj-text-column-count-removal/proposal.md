## Why

F-03/F-06 retain a direct column count when projected columns is deleted. Explicit single-column text and absence must remain distinct while spacing and direction stay unchanged.

## What Changes

- Complete columns deletion/restoration for the existing 1..16 range across supported text owners and table cells.
- Extend simple-style removal under existing authority and other-field guards.
- Reuse the numeric lifecycle experiment and update public descriptions and preview limits.

## Capabilities

### New Capabilities
- `ppj-text-column-count-removal`: source-bound columns presence lifecycle.

### Modified Capabilities

None.

## Impact

PPJ source compiler, bounded body admission, shared tests, schema, Help, registry, text reference and F-03/F-06 coverage. Existing NoColumns wire marker is sufficient.
