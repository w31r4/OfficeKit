# Design

## Public state

Native bar, column, pie and doughnut series may declare:

```json
"pointStyles": [
  { "index": 2, "fill": "#D9A21B" },
  {
    "index": 5,
    "fill": { "type": "solid", "color": "#16324F", "opacity": 0.82 },
    "stroke": { "color": "#FFFFFF", "width": 1 },
    "explosion": 12
  }
]
```

Indexes are zero-based, strictly increasing and unique. They must address an
existing non-missing value. Every entry must change at least one of `fill`,
`stroke` or `explosion`. `explosion` is legal only for pie and doughnut and is
bounded to native 0..400 percent.

The initial profile deliberately excludes line/scatter/radar marker overrides,
bubble sizing, picture fills and arbitrary point metadata. Those require
different visual semantics rather than a misleading shared field.

## Native mapping

- One `pointStyles[]` entry maps to one ordered `c:dPt`.
- `index` maps to the required `c:idx`.
- `fill` and `stroke` map to a canonical `c:spPr` using the existing bounded
  direct fill and line profiles.
- `explosion` maps to `c:explosion` on pie/doughnut points.
- Authored nodes use canonical child order and omit empty containers.

The reader accepts only `c:idx`, optional `c:spPr`, optional legal
`c:explosion`, and no attributes. Marker, invert-if-negative, bubble3D,
pictureOptions, extensions, effects, unknown children or irregular paint/line
graphs make the containing native chart non-editable while preserving bytes.

## Source continuation

Recognized charts continue to use the existing `setChartFill` capability.
The PPJ differ may add, replace or remove only bounded point-style entries;
series identity, point count and missing-value topology remain fixed. The
codec patches the existing ChartPart and proves the requested semantics by a
second import.
