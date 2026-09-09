## Why

F-04 connector arrow dimensions exist in the native codec and issued leaves but are absent from PPJ authored fields and fresh projection. Arrow shape edits now work; dimensions need the same complete language lifecycle.

## What Changes

- Add startArrowWidth/startArrowLength/endArrowWidth/endArrowLength with sm/med/lg values.
- Preserve optional attribute presence through authored export and fresh projection, and support source-bound size changes/removal through setConnectorArrows.
- Keep deleted-arrow dimension cleanup and reject orphan size authoring or explicit size edits without an arrow.

## Capabilities

### New Capabilities
- `ppj-connector-arrow-sizes`: Explicit connector arrow width/length fields with source-bound lifecycle.

### Modified Capabilities
None; prior connector changes remain unarchived with no owning main spec.

## Impact

PPJ schema, authored/projected/source compiler paths, existing capability vocabulary, focused native regression, Help/registry/reference docs. Existing optional native strings suffice; no wire change or new host-routing claim.
