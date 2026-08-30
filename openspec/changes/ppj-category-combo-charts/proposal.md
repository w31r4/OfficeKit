# PPJ category combo charts

## Why

PPJ advertises per-series `area` and `scatter` overrides inside `combo`, but
the native compiler currently owns only one bar/line topology. The accepted
JSON therefore overstates the executable language while hiding a common
category-axis area overlay.

## What changes

- Make the public combo vocabulary truthful: column, line and area series on a
  shared categorical domain.
- Support two or three distinct native plot families with one optional
  secondary category/value axis pair.
- Compile, project, edit and reimport this bounded topology as one native
  ChartPart.
- Remove the unimplemented scatter override from the combo-series schema.

## What does not change

- Numeric scatter/bubble overlays require a future numeric-axis combo contract.
- Candlestick and vector-lowered chart families do not become ChartPart series.
- Horizontal bars, duplicate plot families split across axis pairs and
  unrelated arbitrary OOXML combo graphs remain rejected or source-preserved.

