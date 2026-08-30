# Context

Kimi exposes a symbol on bar series. OfficeKit already owns two bounded vector
catalogs: named offline icons and DrawingML preset geometries. PPJ should reuse
those catalogs rather than introduce a third symbol language.

# Decisions

## 1. Reuse bar and column data

A pictographic chart remains `chartType: "bar"` or `"column"`. Exactly one
series declares:

```json
{
  "symbol": {
    "kind": "icon",
    "iconName": "fas:user",
    "unit": 10,
    "gap": 2,
    "showValue": true
  }
}
```

`kind: "preset"` uses an existing preset geometry name instead. Series color,
fill and stroke remain the paint vocabulary.

## 2. Keep expansion bounded and exact

The profile accepts one series, 2..12 categories, non-negative finite values,
and values that are exact multiples of a positive `unit`. It emits at most 32
symbols per category and 192 symbols in total. A non-divisible value fails
closed; the compiler does not crop a glyph or imply false precision.

Horizontal bars lay symbols left to right. Columns stack symbols bottom to top.
The compiler adds stable category, value and unit labels as native text.

## 3. Lower to editable DrawingML

Named icons reuse `PpjIconCatalog`; preset symbols reuse
`PptxPresetGeometryAdjustmentCodec`. Every visible unit is an editable native
shape with a stable child ID. The chart element itself becomes one editable
group, never an image or opaque SVG.

## 4. Recovery stays honest

An OfficeKit-authored PPTX restores exact pictographic semantics from its
embedded PPJ. If the snapshot is removed, the native group projects as editable
shapes. OfficeKit does not guess data or symbol units from arbitrary groups.

# Lean verification

Extend the existing comprehensive authored-PPJ contract with one pictographic
chart and one invalid divisibility case. Assert symbol count, geometry,
determinism, exact recovery and honest snapshot-free group projection. Do not
create a fixture or chart matrix.
