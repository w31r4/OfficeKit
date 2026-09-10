## Why

F-03 still lacks complete source-bound space-after editing. The point and multiplier fields already author and import but cannot independently switch units, delete or restore the native slot through PPJ.

## What Changes

- Complete `text.paragraphs[].style.spaceAfter` and `spaceAfterMultiplier`, including zero, unit switching, deletion/restoration and spacing-only style-wrapper removal.
- Use native 0..1584pt or 0..132 multiplier bounds and ties-to-even rounding. Higher-priority paragraph styles choose both unit and value.
- **BREAKING**: Reject both unit forms in the same style and align schema bounds with the native range.
- Preserve space before, line spacing, native spelling, other paragraph content and unmodeled source spacing during unrelated edits. Unmodeled replacement rejects without output.
- Reuse the space-before lifecycle experiment for both slots, and synchronize Help, registry, references, preview diagnostics and backlog.

## Capabilities

### New Capabilities

- `ppj-paragraph-space-after-lifecycle`: complete direct space-after lifecycle in either native unit.

### Modified Capabilities

None.

## Impact

PPJ schema/compiler/authority, native paragraph-spacing read/write, shared lifecycle fixtures and presentation metadata. Existing wire deletion intent is reused; no wire-version or dependency change. NativeAOT and host layout remain separate evidence.
