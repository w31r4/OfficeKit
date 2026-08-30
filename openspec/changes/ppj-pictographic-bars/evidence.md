# Evidence

## Contract and compiler

- PPJ schema accepts one bounded `series.symbol` for `bar` or `column`.
- Named symbols resolve the pinned Font Awesome Free catalog; preset symbols
  resolve the shared DrawingML preset-geometry catalog.
- NativeAOT emits one editable shape per exact unit plus stable title, unit,
  category and value text.
- Values that require a fractional symbol, exceed 32 units in one category, or
  exceed 192 units in total fail before output.
- Embedded PPJ restores exact symbol semantics. With the snapshot removed, the
  output projects truthfully as an editable group rather than a guessed chart.

## Focused verification

```text
dotnet build native/OfficeKit/src/OfficeKit.Codec/OfficeKit.Codec.csproj \
  -p:UseSharedCompilation=false --no-restore
Build succeeded. 0 warnings, 0 errors.

dotnet test native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj \
  -p:UseSharedCompilation=false --no-restore \
  --filter FullyQualifiedName~PpjV1CompilesCanonicalPresentationProgramDeterministically
Passed: 1, Failed: 0, Skipped: 0.

node skills/presentations/skills/presentation-skill-maintainer/scripts/maintain-presentation-skill.mjs check
Presentation Skill maintenance check ok · 151 Help APIs · 73 native leaves · 13 host-only operations

npx openspec validate ppj-pictographic-bars --strict
Change 'ppj-pictographic-bars' is valid.
```

The first focused-test run exposed a stale assertion that treated projected
structured text as a scalar string. The assertion now reads its paragraph/run;
the compiler output and truthful editable projection were unchanged.

## Boundaries

- Full `npm test`, package/release gates and Windows PowerPoint were not run for
  this bounded primitive slice.
- Fractional icons, multiple series, axes, `chartBuild`, network icons and
  inference from arbitrary third-party groups remain unsupported by design.

## Integration

- The four implementation commits were pushed normally and fast-forwarded to
  `main` without squash, rebase or force push.
