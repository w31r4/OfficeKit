# Add paragraph default reflection start-position leaf

## Why

The PPJ reflection object already carries `startPosition`, and authored
paragraph defaults can write it, but an imported paragraph default reflection
does not expose that native value as an editable leaf. That leaves one direct
DrawingML field in the documented P0 text-effect profile without a complete
source-bound lifecycle.

## What Changes

- Expose `paragraph.style.defaultText.reflection.startPosition` as the
  `textDefaultReflectionStartPosition` native leaf.
- Prove an existing canonical `reflection/@stPos` token on the owning
  paragraph and splice only that token during source-bound edits.
- Keep the profile bounded to a single direct reflection, with a canonical
  `endPos` of `100000`; malformed, unknown, or otherwise unsupported graphs
  remain source-owned or fail closed.
- Add a focused authored → source-bound edit → re-projection regression and
  update the PPJ capability evidence.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-default-reflection-start-position`: source-bound
  projection and token editing for one paragraph default reflection position.

### Modified Capabilities

<!-- No existing main spec is present; this change adds a bounded capability. -->

## Impact

The presentation native leaf projection and edit-plan proof/patch path gain one
numeric leaf. The existing protobuf reflection field and PPJ schema already
represent the semantic value, so no wire-version or relationship changes are
needed.
