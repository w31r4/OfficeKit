## Why

F-03 still lacks the complete PPJ lifecycle of paragraph left indent. `text.paragraphs[].style.indent` authors/imports native `marL`, but ordinary source-bound text/shape owners cannot independently assign, delete or restore it.

## What Changes

- Complete direct paragraph `indent`, including explicit zero, assignment, removal/restoration and indent-only style-wrapper removal.
- **BREAKING**: Align the field with its native nonnegative range, 0..4032pt, and retain nearest-EMU rounding with ties to even.
- Require exact field authority and patch only changed paragraphs, preserving hanging indent, spacing, source numeric spelling, runs, neighbors and non-target XML/ZIP.
- Preserve invalid native left margins during no-op and unrelated scalar edits; reject their replacement without output.
- Add one focused lifecycle fixture and synchronize Help, registry, references, preview diagnostics and backlog.

## Capabilities

### New Capabilities

- `ppj-paragraph-indent-lifecycle`: complete direct paragraph left-indent lifecycle.

### Modified Capabilities

None.

## Impact

PPJ schema/authority/lowering, native paragraph-layout read/write and focused fixtures. Existing `no_margin_left` wire intent is reused. Hanging lifecycle, table/placeholder inheritance, NativeAOT and host layout retain separate boundaries.
