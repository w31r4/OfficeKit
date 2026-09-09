- [x] Add the optional rounded-corners field to PPJ schema, wire model, validator, compiler/projector, registry, and reference docs.
- [x] Extend shared ChartML read/build/patch paths for ordinary and combo chart parts with presence-aware boolean validation.
- [x] Add a focused authored/source-bound/reprojection test and run protocol, OpenSpec, and narrow codec gates.
- [x] Preserve explicit false versus absence through semantic change detection; test false removal and re-addition on ordinary and combo charts using native XML projection.

Validation: six focused codec cases passed with .NET SDK 8.0.417 (installed locally; repository SDK pin unchanged). Ordinary/combo cycles check native XML, independent PPJ projection and byte equality of every unrelated package part. Invalid-value/duplicate/extra-content cases retain their source ChartPart without setChartPlot. OpenSpec strict validation and git diff --check passed.
