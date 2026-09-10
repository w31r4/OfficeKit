## Why

F-03 still lacks complete source-bound line-spacing editing. The point and multiplier fields author and import, but PPJ cannot independently switch their units, remove or restore the native slot.

## What Changes

- Complete `text.paragraphs[].style.lineSpacing` and `lineSpacingMultiplier`, including unit switching, removal/restoration and spacing-only style-wrapper removal.
- Retain positive input and positive rounded native values. Use points up to 1584 or multipliers up to 132, with ties-to-even rounding; values that round to zero reject.
- **BREAKING**: Reject simultaneous unit declarations and align maximum values with the native limits. Higher-priority styles select both unit and value.
- Preserve before/after spacing, native spelling, other paragraph content and unmodeled line spacing during unrelated edits; unmodeled replacement rejects without output.
- Extend the shared spacing lifecycle experiment and synchronize Help, registry, references, preview diagnostics and backlog.

## Capabilities

### New Capabilities

- `ppj-paragraph-line-spacing-lifecycle`: complete direct line-spacing lifecycle in either native unit.

### Modified Capabilities

None.

## Impact

PPJ schema/compiler/authority, native paragraph-spacing read/write and shared fixtures. The existing no-line-spacing wire intent is reused; no protocol or dependency change. NativeAOT and host layout retain separate evidence.
