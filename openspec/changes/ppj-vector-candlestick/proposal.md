# PPJ vector candlestick

## Why

Financial presentations frequently need OHLC or HLC price evidence, but PPJ
currently forces an Agent to approximate it with unrelated shapes or insert a
raster chart. PPTD exposes a candlestick name, yet a long enum is not enough:
OfficeKit needs one finite semantic contract, native editable output, exact PPJ
recovery and honest fallback projection.

## What changes

- Add `candlestick` to the PPJ chart vocabulary.
- Add aligned open/high/low/close series channels and a bounded candlestick
  style.
- Compile OHLC bodies or HLC close ticks, wicks, axes, labels and optional close
  values into one editable DrawingML group.
- Restore exact semantics from embedded PPJ; project the native fallback as an
  ordinary group rather than guessing a chart.
- Document when the primitive is appropriate and how its finite profile differs
  from a native ChartPart.

## Impact

The JSON Schema, C# PPJ parser/validator/compiler, generated PPJ manual,
financial chart guidance, one existing comprehensive native contract and
coverage ledger change. The Office wire version remains 2.
