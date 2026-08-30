# Design

## Public state

An ordinary `chartAxis` may declare:

```json
"axisLineArrow": {
  "start": "none|triangle|stealth|diamond|oval|open",
  "end": "none|triangle|stealth|diamond|oval|open"
}
```

At least one endpoint is required. Omission preserves the default/no-arrow
state. `none` is explicit and lets a source-bound edit remove one recognized
endpoint without changing the axis or chart topology.

## Native mapping

- `start` maps to `c:*Ax/c:spPr/a:ln/a:headEnd`.
- `end` maps to `c:*Ax/c:spPr/a:ln/a:tailEnd`.
- PPJ `open` maps to DrawingML `arrow`; the other names map directly.
- Endpoint width/length are intentionally omitted. Imported custom endpoint
  sizing, unknown children, effects or non-canonical line topology remain
  read-only and byte-preserved.

The line message carries the two endpoint values because it already owns the
axis `a:ln`; validation rejects them on series, marker, trend, error-bar and
grid lines. `spokeAxis` does not accept arrowheads because its category-axis
major gridlines are radial spokes rather than a directed baseline.

## Source continuation

Canonical axes project `axisLineArrow` and retain the existing `setChartAxis`
capability. A bounded edit rebuilds only the recognized axis line container,
then the codec reimports and compares the requested semantics. Adding/removing
an endpoint is allowed only inside that already editable canonical container.
