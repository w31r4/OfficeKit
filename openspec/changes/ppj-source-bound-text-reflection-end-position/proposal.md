# Add direct text reflection end-position leaf

## Why

PPJ text reflections already carry `endPosition`, and authored text can write
it, but an imported direct rich-text run has no editable leaf for the native
endpoint. A one-sided reflection ramp therefore remains source-owned even
when its end token is explicit and independently safe to patch.

## What Changes

- Expose `run.style.reflection.endPosition` as the
  `textReflectionEndPosition` native leaf.
- Prove an existing canonical `a:reflection/@endPos` token on the owning text
  run and splice only that token during source-bound edits.
- Keep the profile narrow: the run reflection must have no variable
  `stPos` (or an explicit canonical `0`); reflections with a variable start,
  two variable endpoints, or unsupported effect topology remain source-owned.
- Add a focused source-bound edit and re-projection regression and update PPJ
  capability evidence.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-reflection-end-position`: source-bound projection and
  token editing for one direct rich-text reflection endpoint.

### Modified Capabilities

<!-- No existing main spec is present; this change adds a bounded capability. -->

## Impact

The direct text reflection projection and edit-plan proof/patch path gain one
numeric leaf. The existing protobuf reflection message and PPJ schema already
represent the semantic value, so no wire-version or relationship changes are
needed.
