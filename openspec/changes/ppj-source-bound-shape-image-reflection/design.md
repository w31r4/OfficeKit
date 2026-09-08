# Design: source-bound shape and image reflection

## Owners

The shape owner is a direct DrawingML shape-properties effect list:

```text
p:sp/p:spPr/a:effectLst/[a:outerShdw/]a:reflection
```

The picture owner is the corresponding direct picture-properties effect list:

```text
p:pic/p:spPr/a:effectLst/[a:outerShdw/]a:reflection
```

The strict reader accepts one direct `a:reflection` with bounded `blurRad`,
`stA`, `endA`, `dist`, and `dir` attributes, deterministic full-span
positions (`stPos=0`, `endPos=100000`), no child elements, and an optional
valid direct outer shadow before it. The preceding outer shadow remains
source-preserved when the reflection owner is edited.

The PPJ projection uses `style.reflection` for ordinary shapes and lines, and
`reflection` for pictures. Source-bound native leaves are owner-specific:

- `shapeReflectionBlurRadiusEmu`,
  `shapeReflectionStartOpacityThousandthPercent`,
  `shapeReflectionEndOpacityThousandthPercent`,
  `shapeReflectionDistanceEmu`, and `shapeReflectionDirectionDegrees`;
- the corresponding `imageReflection...` leaves for pictures.

Only existing native attributes receive leaves. The profile has no color leaf;
reflection opacity is represented by the start/end opacity attributes.

## Safety boundary

Multiple effect lists, reflection without full-span positions, reflection
combined with glow, inner shadow, soft edge, or any other unmodeled effect,
reflection transform variants, unknown children, extensions, missing or
out-of-range geometry/opacity, and malformed shape/picture owners remain
source-owned or fail closed. No source-bound effect insertion, removal,
reordering, or conversion is allowed.

## Edit proof

The native leaf retains the existing source/native hashes and shape-tree
binding. The edit plan re-proves the selected `p:sp` or `p:pic`, direct
`p:spPr/a:effectLst/a:reflection`, full-span topology, and scalar precondition
before token-splicing only the requested attribute in the owning SlidePart.
The focused fixture removes the embedded PPJ snapshot, edits shape and
picture reflection leaves, verifies SlidePart-only mutation and package
validity, and reprojects both values. A non-full-span or complex effect list
must not receive reflection leaves.
