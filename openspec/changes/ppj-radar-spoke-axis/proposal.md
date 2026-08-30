# Change: Add a semantic radar spoke axis

## Why

PPJ can already author editable native standard radar charts and can style their
generic category/value axes. That representation makes an Agent translate a
single visual intent into two unrelated axis objects: radial spokes belong to
category-axis gridlines, while the concentric spider grid and numeric scale
belong to the value axis. The file format should expose the radar concept the
author is actually choosing.

## What changes

- Add one radar-only `spokeAxis` object for visibility, numeric scale, tick
  labels, radial spoke lines and concentric grid lines.
- Add presence-aware tick-label visibility to the bounded native chart-axis
  wire so `label: false` has an editable DrawingML representation.
- Lower `spokeAxis` to the existing paired category/value ChartPart axes and
  project that canonical topology back to the semantic PPJ object.
- Keep generic `xAxis`/`yAxis` compatible, but reject mixing them with
  `spokeAxis` in one radar element.
- Keep alternate radar styles, unsupported tick-label positions and unknown
  native axis graphs source-owned.

## Impact

- PPJ schema and generated language reference
- NativeAOT chart axis reader/writer, PPJ validator/compiler/projector
- Existing comprehensive authored PPJ round-trip contract
- Presentation capability registry and focused chart guidance

