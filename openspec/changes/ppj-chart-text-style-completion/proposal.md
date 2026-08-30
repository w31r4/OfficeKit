# PPJ chart text-style completion

## Why

PPJ already owns chart-title and axis tick-label typography, but legend,
data-label, and axis-title text still fall back to host defaults. These are
visible authoring controls, not obscure OOXML coverage: without them an Agent
cannot establish a coherent type hierarchy across a data page.

## What Changes

- Add `legendTextStyle` to chart style.
- Add `textStyle` to structured data labels.
- Add `titleTextStyle` to each chart axis.
- Reuse the bounded chart text profile for authored output, import projection,
  source-bound editing, semantic hashing, and PPJ documentation.
- Keep effect-bearing, multi-run, per-point, or otherwise irregular native text
  graphs source-owned and fail closed.

## Impact

- Additive wire-v2 fields only; the protocol version remains unchanged.
- Ordinary and combo Presentation charts share the implementation.
- The shared native chart codec also preserves the new fields for worksheet
  charts, without adding a new public Spreadsheet API in this change.
