# Design: source-bound shape and image glow

## Owners

The shape owner is a direct DrawingML shape properties effect list:

```text
p:sp/p:spPr/a:effectLst/a:glow
```

The picture owner is the corresponding direct picture properties effect list:

```text
p:pic/p:spPr/a:effectLst/a:glow
```

The strict reader accepts one direct `a:glow` with a bounded `rad` and one
direct RGB or theme color, with an optional explicit alpha. The effect list
may contain only that glow, or glow followed by one valid direct outer shadow.
The outer shadow remains source-preserved when the glow owner is edited.

The PPJ projection uses `style.glow` for ordinary shapes and lines, and
`glow` for pictures. Source-bound native leaves are owner-specific:

- `shapeGlowRadiusEmu`, `shapeGlowColorRgb` or `shapeGlowColorScheme`, and
  `shapeGlowOpacityThousandthPercent`;
- `imageGlowRadiusEmu`, `imageGlowColorRgb` or `imageGlowColorScheme`, and
  `imageGlowOpacityThousandthPercent`.

The opacity leaf is issued only when the source contains an explicit direct
alpha. The color leaf preserves RGB-versus-theme topology. Radius and color
are always present in a valid direct glow.

## Safety boundary

Multiple effect lists, glow combined with inner shadow, reflection, soft edge,
or any other unmodeled effect, non-direct color transforms, unknown children,
extensions, missing or out-of-range radius/color, and malformed shape/picture
owners remain source-owned or fail closed. No source-bound effect insertion,
removal, reordering, or conversion is allowed.

## Edit proof

The native leaf retains the existing source/native hashes and shape-tree
binding. The edit plan re-proves the selected `p:sp` or `p:pic`, direct
`p:spPr/a:effectLst/a:glow`, strict topology, color kind, and scalar precondition
before token-splicing only the requested attribute in the owning SlidePart.
The focused fixture removes the embedded PPJ snapshot, edits shape and picture
glow leaves, verifies SlidePart-only mutation and package validity, and
reprojects both values. A complex effect list must not receive glow leaves.
