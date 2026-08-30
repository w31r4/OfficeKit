# Context

PowerPoint's native combination-chart topology in OfficeKit is deliberately
bounded to column, line and area plots sharing categorical axes. Scatter and
bubble plots use value axes in both directions, so inserting them into that
writer would produce a structurally misleading chart. Candlestick charts are
already editable vector groups because DrawingML has no faithful native stock
chart surface in the current codec.

# Decisions

## 1. Select numeric combo semantics from the series

`chartType: "combo"` keeps its current native categorical behavior until one
series declares `chartType: "scatter"` or `"bubble"`. That declaration selects
the numeric profile for the whole element:

```json
{
  "type": "chart",
  "chartType": "combo",
  "data": {
    "categories": [],
    "series": [
      {
        "id": "observed",
        "name": "Observed",
        "chartType": "bubble",
        "xValues": [1, 2, 3],
        "values": [4, 7, 8],
        "bubbleSizes": [12, 18, 10]
      },
      {
        "id": "trend",
        "name": "Trend",
        "chartType": "line",
        "xValues": [1, 2, 3],
        "values": [4.5, 6.5, 8.5]
      }
    ]
  }
}
```

The numeric profile accepts 2..8 series and 2..64 complete finite points per
series. Series types are scatter, bubble, line, area and column; at least one
scatter or bubble and one different plot family are required. Every series has
strictly increasing `xValues`. Bubble sizes are positive and appear only on a
bubble series. Shared categories must be empty. Secondary axes, stacking,
trendlines, error bars and formula-backed data are outside this bounded profile.

## 2. Use one explicit value/value plot

The compiler derives one X and one Y domain from all series. Explicit axis
minimum and maximum values may expand or clip neither domain. Area and column
series include zero in the automatic Y domain and use zero as their baseline;
an explicit Y domain that excludes zero therefore rejects those series. Axis
ticks, labels, titles, gridlines and number formats reuse the existing bounded
PPJ axis vocabulary.

Series are painted in deterministic order: area, column, line, scatter, bubble.
This keeps filled marks behind point evidence. Each band, bar, segment, marker,
bubble, axis and label is a native editable object with a stable child ID.

## 3. Extend candlestick without weakening its body contract

The first candlestick series remains the only OHLC/HLC body and keeps the
existing range checks. Zero through four following series may declare line,
area or column. They contain exactly one complete finite value per candlestick
category and do not carry OHLC or numeric-X channels.

The compiler includes all overlay values in the Y domain. Area and column
overlays require zero to be inside the domain. Area and column overlays render
before wicks and bodies; line overlays render after bodies. This makes a moving
average readable and prevents a filled overlay from hiding the price marks.

## 4. Keep recovery honest

The embedded PPJ snapshot restores exact mixed-chart semantics and stable
program IDs. If that snapshot is absent, the emitted DrawingML remains editable
as a group but is not reverse-engineered into a chart program. Imported native
ChartParts continue through the existing typed/source-bound paths.

# Lean verification

Extend the existing comprehensive authored-PPJ contract with one numeric combo
and one candlestick overlay. Assert editable child marks, z-order, exact
snapshot recovery, deterministic output and one invalid coordinate case. Do not
add a chart matrix, fixture deck or benchmark harness.
