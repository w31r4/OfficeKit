# PPJ vector treemap

## Why

PPJ can author ordinary native charts and bounded vector heatmaps/candlesticks,
but cannot yet express hierarchical part-to-whole evidence. PPTD exposes a
treemap keyword but lowers the result to a PNG with private source metadata.
OfficeKit should keep the semantic convenience while producing editable native
PowerPoint objects.

## What changes

- Add `treemap` to the authored PPJ chart vocabulary.
- Add one aligned parent channel and a bounded hierarchy validator.
- Add a deterministic squarified vector lowering with native shapes and text.
- Keep exact embedded-PPJ recovery and truthful group fallback.
- Teach the generated language manual and chart guide when to use the primitive.

## What does not change

- No ChartPart, raster fallback, raw OOXML, recursive PPJ component, or new wire
  operation is introduced.
- Arbitrary imported shape mosaics are not inferred to be treemaps.
