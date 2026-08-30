# Change: Add authored PPJ pictographic bars

## Why

PPJ can author ordinary bar and column charts, but it cannot express repeated
icons or preset shapes as quantitative units. Authors currently have to expand
the symbols by hand, which disconnects the visible marks from the chart data
and makes stable editing unnecessarily verbose.

## What changes

- Add an optional `symbol` to one bar- or column-chart series.
- Reuse OfficeKit's pinned icon and preset-geometry catalogs.
- Compile the bounded profile into editable DrawingML shapes and labels.
- Restore the exact PPJ through the authored-program snapshot.
- Document when pictographic bars clarify a small count and when ordinary bars
  are more truthful.

## What does not change

- No new chart type, wire operation, network icon source or raw SVG field is
  added.
- Fractional symbols are not clipped or approximated.
- Arbitrary imported shape groups are not inferred as pictographic charts.
- Third-party ChartParts do not receive a pictographic mutation capability.
