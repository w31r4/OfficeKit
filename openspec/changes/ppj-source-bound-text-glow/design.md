# Design: source-bound text glow

## Owner

The owner is a direct DrawingML text-run graph:

```text
p:sp/a:txBody/a:p/a:r/a:rPr/a:effectLst/a:glow
```

The reader accepts a direct RGB or bounded theme color, a literal `rad`, and
at most one direct color `alpha`. A direct `a:outerShdw` may follow the glow;
its existing bounded shadow owner is validated but remains a separate owner.

The source-bound PPJ projection exposes the glow under the run's
`textStyle.glow`. The first experiment issues only direct rich-text run leaves:
`textGlowRadiusEmu`, `textGlowColorRgb` or `textGlowColorScheme`, and
`textGlowOpacityThousandthPercent` when the corresponding source scalar is
present. Paragraph `style.defaultText` is read/projected through the existing
text-style owner but does not issue an independent source-bound leaf in this
slice.

## Safety boundary

The effect list must contain exactly one valid glow, or one valid glow followed
by one valid outer shadow. Multiple effect lists, inner shadows, reflections,
soft edges, unknown DrawingML children, extensions, malformed color effects,
and effect graphs on fields/breaks remain source-owned or fail closed. No
source-bound effect insertion, removal, reordering, or conversion is allowed.

## Edit proof

Each issued leaf carries the existing source/native hashes and the exact text
leaf index. The edit plan re-proves the selected `a:r` and direct `a:rPr`
effect graph, then token-splices only the selected scalar attribute in the
owning SlidePart. RGB/theme color changes preserve the selected color topology;
opacity changes require an existing direct `a:alpha`.

The focused fixture removes the embedded PPJ snapshot, projects the PPTX as
source-bound, changes all available glow leaves, checks that only
`ppt/slides/slide1.xml` changed, validates the package, and projects again.
