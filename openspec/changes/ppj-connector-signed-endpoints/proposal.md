## Why

F-04 connector frame anchors can resolve outside a group child origin, but PpjConnectorEndpointResolver and PptxConnectorCodec currently reject every negative endpoint. This rejects valid signed placement and makes the existing rotated cross-group fixture fail.

## What Changes

- Treat connector from/to coordinates as signed points throughout authored export, native import and source-bound edits.
- Validate DrawingML coordinate and extent ranges before arithmetic rather than relying on nonnegativity or unchecked addition.
- Preserve exact endpoint identity, native flip/direction and source no-op/non-target package contents.
- Turn the existing negative-child-space rejection into a positive regression and add a minimal literal negative edit round trip plus malformed-range rejection.

## Capabilities

### New Capabilities
- `ppj-connector-signed-endpoints`: Signed connector endpoints with bounded native coordinates and safe round-trip editing.

### Modified Capabilities
None; the previous connector anchor change remains unarchived and there are no published main specs for it.

## Impact

PPJ endpoint resolver, connector codec and connector native placement proof, existing tests, schema descriptions, capability registry and focused shapes reference. No wire field or production renderer routing change. This continues F-04 without claiming all geometry or host rendering complete.
