## Why

F-03 still leaves DrawingML's `justLow` and `thaiDist` paragraph modes unmodeled. Completing the PPJ alignment enum should preserve these values through authoring, import and source editing instead of leaving valid direct formatting read-only.

## What Changes

- Add PPJ `justifyLow` and `thaiDistributed`, mapping exactly to native `justLow` and `thaiDist`.
- Extend shared paragraph read/write, fresh projection and the bounded table paragraph profile.
- Reuse the seven-value source lifecycle fixture and add a small table/default-owner round-trip check. Retain invalid-token protection and exact field authority.
- Synchronize schema, Help, registry, references, preview diagnostics and backlog; retain partial glyph-layout evidence.

## Capabilities

### New Capabilities

- `ppj-paragraph-alignment-modes`: typed PPJ representation of all seven DrawingML paragraph alignment modes.

### Modified Capabilities

None.

## Impact

PPJ paragraph schema, shared native paragraph codec, table import/edit eligibility, projection, lifecycle fixtures and presentation documentation. Existing wire strings and field authority are reused. Native leaf edits, chart-wide typography and host shaping retain their own contracts.
