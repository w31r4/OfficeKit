## Why

F-03 still lacks a complete PPJ hanging-indent lifecycle. `text.paragraphs[].style.hanging` already authors/imports the negated native paragraph indent, but ordinary source-bound text/shape owners cannot independently assign, remove or restore it.

## What Changes

- Complete hanging indent, including signed values, explicit zero, removal/restoration and hanging-only style-wrapper removal.
- **BREAKING**: Align the field with its native range, -4032..4032pt, retaining nearest-EMU rounding with ties to even.
- Require exact field authority and patch only changed paragraphs, preserving left indent, spacing, numeric spelling, runs, neighbors and non-target XML/ZIP.
- Preserve invalid native indentation during no-op and unrelated scalar edits; reject replacement without output.
- Reuse the left-indent lifecycle fixture and synchronize Help, registry, references, preview diagnostics and backlog.

## Capabilities

### New Capabilities

- `ppj-paragraph-hanging-lifecycle`: complete signed paragraph hanging-indent lifecycle.

### Modified Capabilities

None.

## Impact

PPJ schema/authority/lowering, native paragraph-layout read/write and shared coordinate fixtures. Existing `no_indent` wire intent is reused. Table/placeholder inheritance, NativeAOT and host layout retain separate boundaries.
