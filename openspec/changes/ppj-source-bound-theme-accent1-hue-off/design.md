# Design: source-bound accent1 hueOff

## Boundary

The owner is the canonical shared `ThemePart` selected by the existing PPTX
source-preserving path. Projection and compilation require the same strict
shape:

```text
a:accent1Color
  a:srgbClr val="RRGGBB"
    a:hueOff val="integer"
```

The color owner and RGB leaf have no extra attributes, the RGB value is six
hexadecimal digits, the hueOff node has no descendants or attributes other
than `val`, and the native integer is between -21,600,000 and 21,600,000
inclusive. Any missing, ambiguous, transformed, or extended topology remains
opaque.

## Representation

`PresentationThemeColorTransform.HueOffsetAngleThousandth` already stores the
DrawingML 1/60000-degree integer and the authored compiler already accepts
`hueOff` in -360..360 degrees. The projector exposes the integer divided by
60,000 as a JSON number. The source-bound compiler validates a finite degree
value in that range and uses `Math.Round(value * 60_000d,
MidpointRounding.AwayFromZero)` before assigning the existing authored field.

## Compile path

Add a `hueOffChanged` branch beside the existing accent1 transform branches.
It requires `setThemeAccent1HueOff` and the exact
`accentTransforms.accent1.hueOff` field, requires an imported direct hueOff
owner, and rejects every other transform flag, discrete transform, or changed
theme field. The existing PPTX source writer then changes only the direct
`A.HueOffset.Val` token and records the canonical ThemePart as the changed
part. No-op requests return the original bytes.

## Evidence

The focused test authors a minimal deck, removes the embedded PPJ snapshot,
projects the imported hueOff owner, proves no-op identity and capability
authority, edits 12.5 degrees to 750,000 units, validates the package, checks
all non-theme ZIP parts remain byte-identical, and reprojects 12.5 degrees.
It also covers deletion, an unsupported sibling, a combined theme edit, an
out-of-range value, and capability tampering.
