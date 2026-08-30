# Design

The feature extends `SpreadsheetChartTextStyleArtifact` instead of inventing a
Presentation-only duplicate. The canonical native profile owns optional size,
Latin and East Asian typeface, bold, italic, and one direct `a:solidFill` with
`a:srgbClr` plus optional `a:alpha`. Unknown attributes, theme transforms,
multiple fills, effects, paragraph topology, and other rich-text children keep
the chart outside the editable profile.

PPJ theme tokens are resolved to literal RGB during authored compilation. A
normal PPTX projection can recover executable RGB state, not the original token;
an OfficeKit-authored PPTX still restores its embedded PPJ exactly.

Source-bound edits use one `setChartTextStyle` capability over
`style.titleTextStyle` and each axis `textStyle`. The chart type, title text,
axis topology, and all other style fields must remain equal. The existing chart
part writer patches only the canonical text-property nodes and leaves the slide
part and unrelated chart XML stable.

