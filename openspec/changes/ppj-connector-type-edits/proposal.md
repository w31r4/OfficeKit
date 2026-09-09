## Why

F-04 connectorType authors and projects straight/elbow/curved, while the PPJ source compiler rejects type edits even though the native writer already supports canonical geometry replacement.

## What Changes

- Issue setConnectorType for recognized editable connectors.
- Compile connectorType changes using existing native canonical geometry while retaining endpoints, bindings and line style.
- Verify fresh-source type transitions, no-op bytes, target SlidePart-only differences and rejection outside the issued profile.

## Capabilities

### New Capabilities
- `ppj-connector-type-edits`: Source-bound editing of the existing connectorType field.

### Modified Capabilities
None; no published main spec owns this field lifecycle.

## Impact

PPJ schema capability vocabulary, semantic validator, projector/compiler, focused tests and capability/Help/reference documentation. Existing wire and native geometry writer are reused. Full routing and arbitrary imported geometry remain outside the profile.
