# Source-bound compound shape compositing opacity

## Why

K-04/F-05 still leaves ordinary source-bound shapes without an element-level
opacity owner. Their native alpha may be distributed across a fill, gradient
stops, image paint, outline, and shadow, so projecting one `compositing.opacity`
value without proving those owners agree would silently change the rendering
model.

## What Changes

- Project a bounded `compositing.opacity` field for eligible source-bound
  ordinary shapes with no visible text.
- Issue `setOpacity` for that exact field and lower a requested value to every
  modeled visual paint owner of the shape.
- Keep lines, text-bearing shapes, placeholders, opaque image-fill topology,
  and shapes with mixed owner alpha source-bound.
- Add a focused source-bound PPTX round-trip experiment and record the bounded
  status in coverage/backlog documentation.

## Capabilities

### Modified Capabilities

- `ppj-source-shape-compositing-opacity`: Source-bound ordinary shapes may
  expose a single compound opacity only when all modeled visual owners have
  the same effective alpha.

## Impact

- Native source-bound projector/capability compiler and one native test.
- No protobuf or Office wire-version change; the existing PPJ `compositing`
  contract is reused.
- Blend modes, isolation, clip stacks, group/layer alpha, text paint, and
  unsupported source image-fill graphs remain outside this slice.
