# Design: source-bound text reflection

## Owner

The owner is a direct DrawingML rich-text run:

```text
p:sp/a:txBody/a:p/a:r/a:rPr/a:effectLst/a:reflection
```

The strict reader accepts one direct `a:reflection`, optionally preceded by
one already-valid direct `a:outerShdw`. The reflection must use the canonical
full-span positions `stPos=0` and `endPos=100000`, and may contain only the
modeled scalar attributes `blurRad`, `stA`, `endA`, `dist`, and `dir`.

Only ordinary text runs issue source-bound leaves:

- `textReflectionBlurRadiusEmu`
- `textReflectionStartOpacityThousandthPercent`
- `textReflectionEndOpacityThousandthPercent`
- `textReflectionDistanceEmu`
- `textReflectionDirectionDegrees`

The positions are owner-proof constraints rather than independently exposed
leaves. Paragraph `defaultText` is projected through the existing text-style
owner but does not issue an independent source-bound leaf in this slice.

## Safety boundary

Multiple effect lists, non-full-span positions, reflection transform
attributes, glow, inner shadow, soft edge, unknown DrawingML children,
extensions, malformed scalar values, fields, breaks, and other inline
topologies remain source-owned or fail closed. No source-bound insertion,
removal, reordering, or effect conversion is allowed.

## Edit proof

Each leaf carries the existing source/native hashes and exact text-leaf index.
The edit plan re-proves the selected `a:r`, its direct `a:rPr`, and the strict
reflection graph before token-splicing only the selected scalar attribute in
the owning SlidePart. The focused fixture removes the embedded PPJ snapshot,
changes all five reflection leaves, verifies that only
`ppt/slides/slide1.xml` changed, validates the package, and reprojects the
values.
