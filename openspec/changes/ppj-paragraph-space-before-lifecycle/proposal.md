## Why

F-03 paragraph spacing already authors and imports, but ordinary source-bound paragraphs cannot independently set, remove or restore their space before. Its point and multiplier forms represent the same native slot and need one complete lifecycle.

## What Changes

- Complete `text.paragraphs[].style.spaceBefore` and `spaceBeforeMultiplier`, including unit switching, explicit zero, removal/restoration and spacing-only style-wrapper removal.
- Use 0..1584pt with native hundredths or 0..132 multipliers with native 1/100000 precision, rounded with ties to even. Higher-priority paragraph styles choose the unit as well as its value.
- **BREAKING**: Reject both unit fields in the same paragraph style and align the schema bounds with the existing native range instead of silently choosing a unit or accepting an unusable value.
- Preserve other spacing slots, native spelling, runs and neighboring paragraphs. Unmodeled space-before content remains source-owned during unrelated edits and rejects replacement.
- Synchronize authority, Help, registry, text guidance, generated references and backlog, with explicit partial preview evidence.

## Capabilities

### New Capabilities

- `ppj-paragraph-space-before-lifecycle`: complete direct space-before lifecycle in either unit.

### Modified Capabilities

None.

## Impact

PPJ schema/validation, authored and source-bound compilers, projection authority, native paragraph-spacing read/write, focused lifecycle tests and presentation metadata. Existing protobuf deletion intent is reused; no wire version change or NativeAOT/host acceptance claim.
