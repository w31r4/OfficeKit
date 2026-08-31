# Design

## Public state

A bubble chart or numeric combo containing bubbles may declare:

```json
"style": {
  "bubbleSizeScale": "log",
  "bubbleRadiusRange": [5, 24]
}
```

`bubbleSizeScale` is `sqrt`, `linear` or `log`. `sqrt` makes visible area track
the measured size, `linear` makes radius track it, and `log` compresses
order-of-magnitude ranges. The default remains the existing
`bubbleSizeMode`: `area` maps to `sqrt`, `width` maps to `linear`.

`bubbleRadiusRange` contains exactly two point values, each from 2 through 72,
with the maximum strictly greater than the minimum. The range describes radius,
not diameter. The compiler uses the minimum and maximum across all bubble
series, so equal values remain visually comparable across series. When every
size is equal, every bubble uses the midpoint radius.

Declaring either field selects the bounded vector numeric-chart lowering. This
is explicit because native ChartML cannot promise exact visible radii or a
logarithmic size transform. Without either field, the existing native bubble
ChartPart behavior is unchanged.

## Lowering

- X and Y remain value axes with explicit finite bounds and tick budgets.
- Each bubble becomes one editable ellipse whose stable ID includes chart,
  series and point identity.
- The shared size domain is transformed by `sqrt(value)`, `value`, or
  `log(value)` and linearly mapped into the requested radius range.
- Existing series fill/stroke, title, legend, axis and accessibility state are
  preserved by the bounded vector chart compiler.
- Whole-object animation remains available; native ChartPart build animation
  is unavailable because the output is an editable group.

## Boundaries

Sizes remain finite and strictly positive. Per-series ranges, negative or zero
sizes, formulas, secondary axes, arbitrary statistical transforms and reverse
inference from third-party DrawingML groups fail closed.
