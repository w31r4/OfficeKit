## Implementation evidence

- PPJ `chartStyle` accepts `startAngle` from 0 through 360 for pie/doughnut
  and `holeSize` from 10 through 90 for doughnut.
- Additive protocol-v2 chart fields preserve scalar presence without changing
  the Office wire version.
- The shared ChartSpace codec reads, validates, authors and patches native
  `c:firstSliceAng` and `c:holeSize`; other chart families reject them.
- PPTX projection restores both values and issues `setChartPlot` only for an
  editable pie or doughnut.
- The source-bound compiler requires that capability and patches the existing
  ChartPart without changing series, relationships or chart topology.

## Lean verification

- `PpjV1CompilesCanonicalPresentationProgramDeterministically`: passed. The
  existing comprehensive case now proves authored `135 / 68`, native XML,
  import and projection, capability issuance, a source-bound change to
  `210 / 74`, one changed ChartPart and successful reprojection.
- `npm run proto:check`: passed; generated protocol output is synchronized.
- Presentation Skill maintainer: passed with `151` Help APIs, `73` native
  leaves and `13` host-only operations.
- `npx openspec validate ppj-circular-chart-geometry --strict`: passed.
- `git diff --check`: passed.

No chart-option matrix, new test file, screenshot snapshot, full `npm test`,
package/release gate, Keynote run or Windows PowerPoint playback check was run
for this bounded plot-property slice.
