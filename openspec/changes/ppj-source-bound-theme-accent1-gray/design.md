# Design: source-bound accent1 gray

## Boundary

The owner is the canonical shared `ThemePart` selected by the existing PPTX
source-preserving path. Projection and compilation require the same strict
shape:

```text
a:accent1Color
  a:srgbClr val="RRGGBB"
    a:gray
```

The color owner and RGB leaf have no extra attributes, the RGB value is six
hexadecimal digits, and the gray node has no attributes or descendants. Any
missing, ambiguous, alpha-bearing, transformed, or extended topology remains
opaque. A source with no direct gray child does not receive this capability.

## Representation

`PresentationThemeColorTransform.Gray` already carries the native boolean.
Projection emits `gray: true` for the exact imported owner. The source-bound
schema accepts a boolean so an authorized edit can set it to `false`; authored
theme compilation keeps the existing positive-only flag contract.

## Compile path

Add a `grayChanged` branch beside the existing accent1 transform branches. It
requires `setThemeAccent1Gray` and the exact
`accentTransforms.accent1.gray` field, requires the imported direct gray owner,
and rejects every other transform flag, discrete transform, or changed theme
field. `true` leaves the source bytes unchanged; `false` removes only the
existing direct `A.Gray` node and records the canonical ThemePart as changed.
The source writer never creates a missing gray node.

## Evidence

The focused test authors a minimal deck, removes the embedded PPJ snapshot,
projects the imported gray owner, proves no-op identity and capability
authority, removes gray, validates the package, checks all non-theme ZIP parts
remain byte-identical, and reprojects without the transform. It also covers
deletion, an unsupported sibling, a combined theme edit, an invalid type, and
capability tampering.
