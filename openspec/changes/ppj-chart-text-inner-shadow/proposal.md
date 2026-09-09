## Why

F-07 still cannot express a chart character inner shadow even though ordinary PPJ text and the wire already have the native effect. Add the field through the shared chart text style so native chart edits preserve independent effects and their optional attributes.

## What Changes

- Add chartTextStyle.innerShadow with required color and optional blur, distance, angle and opacity; preserve omission and explicit zero.
- Support ordered glow, inner shadow, outer shadow and soft edge on recognized chart character owners, including rich trendline labels and vector defaults/overrides.
- Prove a bounded authored/source-bound/reprojection lifecycle and reject unknown or malformed effect graphs; synchronize protocol, references and backlog.

## Capabilities

### New Capabilities
- `ppj-chart-text-inner-shadow`: typed chart character inner-shadow state and independent lifecycle.

### Modified Capabilities
None.

## Impact

PPJ schema, SpreadsheetChartTextStyleArtifact and generated wire, shared DrawingML effect codec, chart style parser/writer, authored/source-bound compilation, projection, vector defaults, focused native/JS regressions and capability/Skill documentation. No production preview switch or host acceptance claim; F-07 remains open.
