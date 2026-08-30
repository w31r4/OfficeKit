# Design

## Language mapping

PPJ keeps its existing color spelling:

```json
{ "color": "#16324FCC" }
```

or a named design color with alpha:

```json
{ "color": { "token": "muted", "alpha": 0.72 } }
```

The compiler resolves the authored color to RGB plus opacity. Opaque colors
omit the native alpha node; translucent colors emit one canonical `a:alpha`
value in thousandths of a percent.

## Native boundary

Two additive optional wire fields carry the alpha for text runs and default run
properties. Readers recognize only a direct RGB or theme color with zero or one
direct `a:alpha` child. Tint, shade, luminance, gradient, pattern, and multiple
transform graphs remain unmodeled and source-preserved.

Projection emits eight-digit hex for direct RGB and `{ token, alpha }` for a
theme color. Authored design tokens resolve to direct RGB, matching the existing
PPJ compiler boundary.

## Source-bound behavior

Imported PPJ may expose recognized text opacity, but ordinary text replacement
does not authorize changing it. A mutation without a dedicated native leaf or
typed capability continues to fail closed.
