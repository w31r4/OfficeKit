## Why

F-03/F-06 text margins retain old source values when PPJ deletes an edge or the margins object. Zero must remain an explicit inset, distinct from absence.

## What Changes

- Complete removal/restoration of left, top, right and bottom margins across supported text owners and table cells.
- Support whole margins and guarded simple-style removal, preserving other edges and text topology.
- Extend the shared lifecycle experiment and publish field semantics and preview limits.

## Capabilities

### New Capabilities
- `ppj-text-margins-removal`: direct text inset presence lifecycle.

### Modified Capabilities

None.

## Impact

PPJ compiler, bounded body codec admission, lifecycle tests, schema/Help/registry, text reference and gap coverage. Existing native inset deletion markers require no wire change.
