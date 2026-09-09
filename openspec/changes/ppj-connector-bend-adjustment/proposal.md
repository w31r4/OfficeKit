## Why

F-04 connector routing currently accepts only default adj1 geometry. A manually positioned bend cannot be authored in PPJ and a non-default imported bend becomes opaque even when its geometry is otherwise simple.

## What Changes

- Add optional integer bendAdjustment for the existing elbow/curved three-segment geometry families, with explicit zero and omission preserved.
- Add a presence-aware native wire field and recognized direct adj1 literal read/write support.
- Extend source-bound connector geometry authority for independent adjustment edit/removal, preserving endpoints, bindings and non-target parts.
- Keep unmodeled guide graphs and unsupported rotated adjusted topology opaque; preview must report any unpainted adjustment.

## Capabilities

### New Capabilities
- `ppj-connector-bend-adjustment`: Native literal bend adjustment with authored and source-bound lifecycle.

### Modified Capabilities
None; prior connector changes remain unarchived.

## Impact

PPJ schema/compiler/projector, connector codec, protobuf bindings, focused connector tests, preview diagnostics and Help/registry/reference documentation. No automatic routing or host appearance claim.
