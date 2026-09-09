## Why

F-03/F-06 compatibleLineSpacing retains booleans but cannot remove its direct source hint. False and absence must remain distinct without altering actual paragraph line spacing.

## What Changes

- Add native deletion intent while preserving the existing optional boolean encoding.
- Complete source removal/restoration and simple-style deletion for supported text owners and tables.
- Reuse boolean/wire experiments with explicit line spacing and update public semantics.

## Capabilities

### New Capabilities
- `ppj-text-compatible-line-spacing-removal`: compatible line-spacing hint presence lifecycle.

### Modified Capabilities

None.

## Impact

Proto/generated JS, native validation/application/normalization, PPJ compiler, shared tests and documentation. New deletion requires updated codec; host line metrics remain outside this increment.
