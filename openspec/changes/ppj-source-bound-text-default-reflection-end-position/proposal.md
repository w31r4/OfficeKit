# Add paragraph default reflection end-position leaf

## Why

The PPJ reflection object already carries `endPosition`, and authored
paragraph defaults can write it, but an imported paragraph default reflection
does not expose the native value as an editable leaf. That leaves the second
endpoint of the bounded reflection ramp outside the source-bound field
surface.

## What Changes

- Expose `paragraph.style.defaultText.reflection.endPosition` as the
  `textDefaultReflectionEndPosition` native leaf.
- Prove an existing canonical `reflection/@endPos` token on the owning
  paragraph and splice only that token during source-bound edits.
- Keep the profile bounded to a direct reflection whose `startPos` is absent
  or the canonical `0`; a reflection with both variable endpoints remains
  source-owned.
- Add a focused authored → source-bound edit → re-projection regression and
  update PPJ capability evidence.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-default-reflection-end-position`: source-bound
  projection and token editing for one paragraph default reflection endpoint.

### Modified Capabilities

<!-- No existing main spec is present; this change adds a bounded capability. -->

## Impact

The presentation native leaf projection and edit-plan proof/patch path gain one
numeric leaf. The existing protobuf reflection field and PPJ schema already
represent the semantic value, so no wire-version or relationship changes are
needed.
