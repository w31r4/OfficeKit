# Add direct text reflection start-position leaf

## Why

The PPJ text reflection object already carries `startPosition`, and authored
text can write it, but an imported direct rich-text run reflection currently
has no editable leaf for that native endpoint. The bounded source-bound text
effect profile therefore cannot preserve and edit a one-sided reflection ramp
without treating the run as opaque.

## What Changes

- Expose `run.style.reflection.startPosition` as the
  `textReflectionStartPosition` native leaf.
- Prove an existing canonical `a:reflection/@stPos` token on the owning text
  run and splice only that token during source-bound edits.
- Keep the profile narrow: the run reflection must have no variable
  `endPos` (or an explicit canonical `100000`); reflections with a variable
  end position or unsupported effect topology remain source-owned.
- Add a focused authored → source-bound edit → re-projection regression and
  update PPJ capability evidence.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-reflection-start-position`: source-bound projection
  and token editing for one direct rich-text reflection endpoint.

### Modified Capabilities

<!-- No existing main spec is present; this change adds a bounded capability. -->

## Impact

The direct text reflection projection and edit-plan proof/patch path gain one
numeric leaf. The existing protobuf reflection message and PPJ schema already
represent the semantic value, so no wire-version or relationship changes are
needed.
