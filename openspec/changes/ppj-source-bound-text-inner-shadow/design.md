# Design: source-bound text inner shadow

## Owner

The owner is a direct DrawingML rich-text run:

```text
p:sp/a:txBody/a:p/a:r/a:rPr/a:effectLst/a:innerShdw
```

The strict reader accepts one direct `a:innerShdw`, optionally followed by
one already-valid direct `a:outerShdw`. The inner shadow may carry direct RGB
or bounded theme color, optional `blurRad`, `dist`, `dir`, and at most one
direct color `alpha`. PPJ projects the result under the run's
`textStyle.innerShadow`.

Only ordinary text runs issue source-bound leaves:

- `textInnerShadowBlurRadiusEmu`
- `textInnerShadowDistanceEmu`
- `textInnerShadowDirectionDegrees`
- `textInnerShadowColorRgb` or `textInnerShadowColorScheme`
- `textInnerShadowOpacityThousandthPercent` when a direct alpha exists

Paragraph `defaultText` is projected through the existing text-style owner but
does not issue an independent source-bound leaf in this slice.

## Safety boundary

Multiple effect lists, glow before inner shadow, reflection, soft edge,
unknown DrawingML children, extensions, malformed color effects, fields,
breaks, and other inline topologies remain source-owned or fail closed. No
source-bound insertion, removal, reordering, or effect conversion is allowed.

## Edit proof

Each leaf carries the existing source/native hashes and exact text-leaf index.
The edit plan re-proves the selected `a:r`, its direct `a:rPr`, and the strict
inner-shadow graph before token-splicing only the selected scalar attribute in
the owning SlidePart. RGB/theme edits preserve the existing color topology;
opacity edits require an existing direct alpha. The focused fixture removes
the embedded PPJ snapshot, changes every available inner-shadow leaf, verifies
that only `ppt/slides/slide1.xml` changed, validates the package, and
reprojects the values.
