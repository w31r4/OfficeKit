# Change: Add truthful PPJ mixed numeric overlays

## Why

PPJ can author editable scatter, bubble and categorical combination charts, and
it can lower a candlestick chart to editable DrawingML. It cannot yet express
two common analytical graphics without manually expanding every mark:

- numeric scatter or bubble evidence overlaid with a line, area or column
  series on the same value/value coordinate system;
- a candlestick body overlaid with moving-average lines, volume columns or a
  bounded area band.

Treating either graphic as an ordinary categorical ChartPart would make the PPJ
program lie about the coordinate system. Flattening it to an image would lose
editability.

## What changes

- Extend `chartType: "combo"` with a bounded numeric profile selected whenever
  any series is scatter or bubble.
- Permit line, area and column series in that profile to carry explicit numeric
  `xValues`.
- Extend `chartType: "candlestick"` with up to four aligned line, area or column
  overlay series after the OHLC/HLC body series.
- Lower both profiles into editable DrawingML groups with stable child IDs,
  explicit numeric scaling and deterministic z-order.
- Preserve exact PPJ through the authored-program snapshot and project a
  snapshot-free result honestly as a native group.

## What does not change

- Existing native categorical combo ChartParts remain unchanged.
- Arbitrary third-party groups are not inferred as semantic mixed charts.
- The compiler does not add a second value axis, chart formulas, raw OOXML,
  automatic statistical transforms or network data.
- Missing coordinates, ambiguous baselines, invalid OHLC ranges and excessive
  mark counts fail closed.
