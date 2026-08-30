# Evidence

## Implemented contract

- Authored PPJ compiles chart-level data-label number formats, bubble scale and
  native area/width size semantics, axis reversal, and bounded direct axis and
  major-grid line styling.
- Projection returns the same semantic fields and issues separate
  `setChartLabels`, `setChartPlot`, and `setChartAxis` capabilities.
- Source-bound edits patch the existing ChartPart, preserve chart topology,
  and recover the requested values after a second PPJ projection.
- Unsupported per-point labels, pixel-radius ranges, logarithmic transforms,
  axis arrows, and irregular native line graphs remain source-owned.

## Lean verification

The implementation used one existing comprehensive PPJ test rather than a new
effect matrix.

- `dotnet build native/OfficeKit/src/OfficeKit.Codec/OfficeKit.Codec.csproj -p:UseSharedCompilation=false --no-restore`: passed with 0 warnings and 0 errors.
- `dotnet test native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj -p:UseSharedCompilation=false --no-restore --filter 'FullyQualifiedName~PpjV1CompilesCanonicalPresentationProgramDeterministically' --logger 'console;verbosity=minimal'`: passed, 1 of 1 focused tests.
- `npm run proto:check`: passed.
- `node skills/presentations/skills/presentation-skill-maintainer/scripts/maintain-presentation-skill.mjs check`: passed with 151 Help APIs, 73 native leaves, and 13 host-only operations.
- `npx openspec validate ppj-native-chart-formatting --strict`: passed.

Not run: full `npm test`, package/release gates, visual rendering, or Windows
PowerPoint. Those remain PPJ 2.0 release-level work rather than evidence needed
for this bounded native-chart formatting slice.
