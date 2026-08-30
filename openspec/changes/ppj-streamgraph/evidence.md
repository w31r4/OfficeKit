# Evidence

## Implemented contract

- `chartType: "area"` plus `style.stacking: "stream"` compiles 2..12
  complete non-negative series over 3..64 ordered categories.
- NativeAOT centers each raw category total and emits one stable closed cubic
  DrawingML path per series, followed by bounded title, category and legend
  text. Solid/gradient fills and bounded direct strokes remain editable.
- Authored PPTX restores the exact PPJ program and stable IDs. Removing the
  embedded program produces a truthful ordinary group projection rather than
  invented stream semantics.
- Non-area use, negative/missing data, empty category totals, secondary axes,
  ChartPart build animation and unsupported series fields fail closed.

## Lean verification

- `dotnet build native/OfficeKit/src/OfficeKit.Codec/OfficeKit.Codec.csproj -p:UseSharedCompilation=false --no-restore`: passed with 0 warnings and 0 errors.
- The existing `PpjV1CompilesCanonicalPresentationProgramDeterministically`
  contract was extended with one three-band streamgraph, one invalid-family
  case, native-path assertions, exact embedded recovery and snapshot-free group
  projection. The focused test passed 1 of 1.
- Its first run exposed a real line-state conflict in legend swatches; the fix
  prevents a swatch from inheriting a series stroke, and the same test then
  passed without adding another fixture.
- `node skills/presentations/skills/presentation-skill-maintainer/scripts/maintain-presentation-skill.mjs check`: passed with 151 Help APIs, 73 native leaves and 13 host-only operations.
- `npx openspec validate ppj-streamgraph --strict`: passed.

Not run: full `npm test`, package/release gates, rendering matrix, or Windows
PowerPoint. Those remain PPJ 2.0 release-level evidence, not prerequisites for
this bounded authored-vector slice.
