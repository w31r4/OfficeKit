# Change: PPJ compound shape opacity

## Why

PPJ already exposes `shape.style.opacity`, but the authored compiler accepts a
value below one only for a solid-fill-only shape. Real presentation overlays,
diagram nodes and visual masks commonly combine a gradient or image fill with
an outline, shadow and explicit text paint. Those valid programs currently fail
at build time and push Agents toward separate overlay objects.

## What changes

- Lower one bounded shape opacity into every directly owned visual branch.
- Multiply existing fill, line, shadow, text, bullet and gradient alpha rather
  than replacing branch-local transparency.
- Support solid, gradient and image fills plus explicit RGB/theme/gradient text.
- Reject only branches whose effective alpha cannot be represented without
  inventing paint, such as inherited text color or alpha-less highlight paint.
- Keep imported effect graphs source-owned.

## Scope

This change does not add a group-compositing engine, flatten a shape to an
image, or reinterpret third-party effect lists. It compiles authored PPJ into
the existing editable DrawingML alpha fields and requires every visible branch
to remain semantically representable.
