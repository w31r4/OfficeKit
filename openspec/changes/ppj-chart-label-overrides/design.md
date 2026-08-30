# Design

## Public state

Any native `chartSeries` may declare:

```json
"dataLabels": {
  "showValue": true,
  "numberFormat": "0.0%",
  "textStyle": { "fontSize": 9, "color": "#334155" },
  "points": [
    { "index": 2, "showValue": false },
    { "index": 5, "numberFormat": "$0.0M", "position": "outside-end" }
  ]
}
```

The fields `showValue`, `showCategory`, `showSeries`, `showPercent`,
`position`, `textStyle`, and `numberFormat` are optional at both levels.
Point indexes are zero-based, unique, strictly increasing in canonical
projection, and must address an existing non-missing series point. Every
point object must override at least one field.

Omission inherits the containing native level. A series object may contain
only `points`; it does not invent false-valued defaults. This presence-aware
state is distinct from the existing plot-level `style.dataLabels`, whose
historical booleans remain compatible.

## Native mapping

- `series[].dataLabels` maps to the corresponding `c:ser/c:dLbls`.
- Series default fields map to direct children of that container.
- Each `points[]` entry maps to one `c:dLbl` with a required `c:idx`.
- Direct number formats always write `sourceLinked=false`.
- Text style reuses the bounded canonical chart text profile.

The reader accepts only this finite child graph. Custom label text (`c:tx`),
manual layout, shape/effect properties, separators, leader-line graphs,
deletion state, extensions, source-linked formats, unknown children, and
non-canonical endpoint order remain outside the profile.

## Source continuation

An otherwise editable chart advertises `setChartLabels`. The PPJ differ may
add, change, or remove a bounded series container or sparse point override,
but cannot change series/point topology. The codec patches the existing
ChartPart, reimports it, and compares the requested semantics. Unrelated OPC
parts and unsupported label graphs remain untouched.
