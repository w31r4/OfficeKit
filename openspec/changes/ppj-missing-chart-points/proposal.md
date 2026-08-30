# PPJ missing chart points

## Why

PPJ already permits `null` inside chart series values, but the authored
compiler rejects it and canonical ChartSpace projection requires a dense
numeric cache. The public language therefore cannot distinguish a missing
observation from numeric zero even though the schema says it can.

## What Changes

- Add bounded missing-value indexes to the existing chart-series wire model.
- Compile PPJ `null` values to a native cache with the full `c:ptCount` and no
  `c:pt` entry at each missing index.
- Reproject that canonical sparse cache as PPJ `null`.
- Keep X values and bubble sizes dense and finite.
- Preserve missing-point topology for source-bound edits; adding or removing a
  missing point remains fail closed.

## Impact

- Additive protobuf field; Office wire remains version 2.
- The PPJ schema is unchanged because nullable values are already declared.
- The shared PPTX/XLSX ChartSpace profile learns one deterministic sparse-Y
  cache representation without accepting arbitrary malformed caches.
