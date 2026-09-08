# Design: source-bound shape and image inner shadow

## Owners

The shape owner is a direct DrawingML shape-properties effect list:

```text
p:sp/p:spPr/a:effectLst/a:innerShdw
```

The picture owner is the corresponding direct picture-properties effect list:

```text
p:pic/p:spPr/a:effectLst/a:innerShdw
```

The strict reader accepts one direct `a:innerShdw` with bounded `blurRad`,
`dist`, and `dir` attributes, one direct RGB or theme color, and an optional
explicit alpha. The effect list may contain only that inner shadow, or the
inner shadow followed by one valid direct outer shadow. The outer shadow
remains source-preserved when the inner-shadow owner is edited.

The PPJ projection uses `style.innerShadow` for ordinary shapes and lines, and
`innerShadow` for pictures. Source-bound native leaves are owner-specific:

- `shapeInnerShadowBlurRadiusEmu`, `shapeInnerShadowDistanceEmu`,
  `shapeInnerShadowDirectionDegrees`, `shapeInnerShadowColorRgb` or
  `shapeInnerShadowColorScheme`, and
  `shapeInnerShadowOpacityThousandthPercent`;
- the corresponding `imageInnerShadow...` leaves for pictures.

The opacity leaf is issued only when the source contains an explicit direct
alpha. The color leaf preserves RGB-versus-theme topology. Geometry leaves are
issued only for existing attributes.

## Safety boundary

Multiple effect lists, inner shadow combined with glow, reflection, soft edge,
or any other unmodeled effect, non-direct color transforms, unknown children,
extensions, missing or out-of-range geometry/color, and malformed
shape/picture owners remain source-owned or fail closed. No source-bound
effect insertion, removal, reordering, or conversion is allowed.

## Edit proof

The native leaf retains the existing source/native hashes and shape-tree
binding. The edit plan re-proves the selected `p:sp` or `p:pic`, direct
`p:spPr/a:effectLst/a:innerShdw`, strict topology, color kind, and scalar
precondition before token-splicing only the requested attribute in the owning
SlidePart. The focused fixture removes the embedded PPJ snapshot, edits shape
and picture inner-shadow leaves, verifies SlidePart-only mutation and package
validity, and reprojects both values. A complex effect list must not receive
inner-shadow leaves.
