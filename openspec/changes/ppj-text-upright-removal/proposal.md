## Why

F-03/F-06 text-body fields need complete source-edit lifecycles. PPJ upright already authors and projects true/false, but omission from an edited source style silently preserves the previous native upright attribute.

## What Changes

- Make removal of a previously projected upright property delete the native attribute, distinct from false.
- Apply the same contract to text/shape/owner-local placeholder styles and structured table-cell text styles; permit removing an otherwise empty upright-only style owner.
- Preserve unrelated body properties, text topology, source bytes on no-op and non-target package members; keep existing capability and unsupported-owner checks.
- Document authored true/false/absence and source add/change/remove with explicit preview limitations.

## Capabilities

### New Capabilities
- `ppj-text-upright-removal`: Complete PPJ text-body upright presence lifecycle.

### Modified Capabilities
None.

## Impact

PPJ source text style merge/table body builder, existing native no_upright operation, focused regression, Help/schema descriptions/registry/text reference and coverage. No new field or wire revision.
