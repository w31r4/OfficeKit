## Why

F-03's paragraph `level` already authors and projects integers 0..8, but ordinary source-bound text/shape owners lack independent field authority and lowering. Native reads and unconditional writes also need to preserve malformed values and unchanged lexical spelling.

## What Changes

- Complete `text.paragraphs[].style.level` assignment, explicit zero, deletion, restoration and level-only style removal.
- Issue and require exact field authority, and patch only changed paragraphs.
- Read native `lvl` as a bounded raw integer; preserve unmodeled values on no-op and unrelated edits while rejecting replacement.
- Add one focused lifecycle fixture, then synchronize Help, schema guidance, registry, references, preview diagnostics and backlog.

## Capabilities

### New Capabilities

- `ppj-paragraph-level-lifecycle`: complete direct paragraph level lifecycle and source preservation.

### Modified Capabilities

None.

## Impact

PPJ source authority/semantic validation/lowering, shared native paragraph properties, focused tests and presentation documentation. The existing optional level field carries deletion through complete paragraph state; wire layout is unchanged. List/master level identifiers and inherited layout retain their existing contracts.
