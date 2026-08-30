# Design

## Public program

```json
{
  "type": "chart",
  "id": "risk-profile",
  "chartType": "radar",
  "frame": { "x": 1, "y": 1.5, "width": 5.5, "height": 4.5 },
  "spokeAxis": {
    "show": true,
    "min": 0,
    "max": 100,
    "majorUnit": 20,
    "label": { "numberFormat": "0", "fontSize": 9, "color": "#475569" },
    "axisLine": { "color": "#CBD5E1", "width": 0.75 },
    "gridLine": { "color": "#E2E8F0", "width": 0.5 }
  },
  "data": {
    "categories": ["Liquidity", "Growth", "Margin", "Resilience"],
    "series": [{ "id": "current", "name": "Current", "values": [72, 81, 64, 77] }]
  }
}
```

`spokeAxis` is optional and radar-only. Once present, omitted `show`, `label`,
`axisLine` and `gridLine` values use the visible radar defaults. `min`, `max`
and `majorUnit` remain optional. A false label hides only numeric tick labels;
the category names around the perimeter remain visible.

## Native lowering

The compiler keeps one native `c:radarChart` and its paired axes:

- `show` controls both `c:catAx/c:delete` and `c:valAx/c:delete`;
- `min`, `max` and `majorUnit` lower to the value axis;
- `label` lowers to value-axis `c:tickLblPos`, `c:numFmt` and `c:txPr`;
- `axisLine` lowers to category-axis major gridlines, which PowerPoint renders
  as the spokes from the center to each category;
- `gridLine` lowers to value-axis major gridlines, which PowerPoint renders as
  the concentric spider grid.

The shared axis wire gains one optional `tick_labels_visible` boolean. False
writes canonical `c:tickLblPos val="none"`; true removes an explicit hidden
position and lets the native default show labels. High/low custom native label
positions stay outside the bounded profile instead of being normalized.

## Projection

Canonical standard-radar axes project to `spokeAxis` when their state fits this
mapping. If a source chart uses axis titles, reversal, category-label styling,
custom axis lines or another unrepresentable axis distinction, projection keeps
the existing explicit `xAxis`/`yAxis` objects rather than discarding state.
Embedded OfficeKit PPJ continues to restore the exact authored object.

## Boundaries

- `spokeAxis` cannot coexist with `xAxis`, `yAxis` or secondary axes.
- `show: false` hides the coordinate system; scale values may remain as dormant
  program state, but line and label visibility compile hidden.
- No log scale, minor grid, filled radar, 3D, theme/effect line graph or raw
  chart XML is introduced.
- Generic chart axes gain only tick-label visibility, not arbitrary native
  label placement.

## Verification

Extend the existing authored PPJ radar object and its package/reprojection
assertions. Add one conflict/invalid-domain assertion to that same contract.
Do not create a radar-style matrix or a new fixture.

