## Why

F-04 still loses native path extrusion eligibility during PPJ projection and path rewrites. Native paths already preserve this optional boolean, so PPJ should carry its presence without implying 3-D rendering.

## What Changes

- Add optional boolean `geometry.paths[].extrusionOk` and preserve true, false and omission.
- Round-trip the field through shared custom-path authoring and projection, including existing source path edits.
- Verify original-source changes/removal and retention during an unrelated path edit; retain explicit preview limitations.

## Capabilities

### New Capabilities
- `ppj-custom-path-extrusion`: Presence-aware path extrusion eligibility in PPJ.

### Modified Capabilities
None.

## Impact

PPJ schema, shared authored/projected custom paths, focused native regression, Help/registry/reference and generated documentation. Existing native wire and path-edit authority suffice.
