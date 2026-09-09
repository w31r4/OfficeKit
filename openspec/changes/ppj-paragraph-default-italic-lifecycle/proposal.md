## Why

F-03 paragraph defaults now expose a complete bold lifecycle, but defaultText.italic edits still fail source-bound style classification. Italic needs the same explicit false and removal semantics.

## What Changes

- Add exact defaultText.italic paragraph-style field authority for ordinary text/shape owners.
- Support true, false, omission, italic-only wrapper removal and restoration.
- Preserve default bold, other defaults, direct runs, unknown native attributes and source ZIP state.
- Share boolean lifecycle tests and extend direct attribute writing to changed bold/italic flags.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-italic-lifecycle`: Direct paragraph default italic presence.

### Modified Capabilities

None.

## Impact

Capability schema/projection/validation, source compiler, default-run codec and presentation guidance. Existing wire fields suffice; other default fields and inherited owners retain their boundaries.
