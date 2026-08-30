# Evidence

## Implemented contract

- `data.series[0].levels` is typed PPJ state for treemap and sunburst only.
- Treemap accepts 1..8 visible levels and makes a truncated branch a summary
  leaf with its full allocated rectangle.
- Sunburst accepts 1..6 visible levels and reallocates the radius across the
  visible rings.
- Complete categories, values and parents remain validated and embedded in PPJ.
  Snapshot-free import exposes only the generated editable group.

## Narrow verification

```text
dotnet build native/OfficeKit/src/OfficeKit.Codec/OfficeKit.Codec.csproj \
  -p:UseSharedCompilation=false --no-restore
```

Passed with zero warnings and zero errors.

```text
dotnet test native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj \
  -p:UseSharedCompilation=false --no-restore \
  --filter FullyQualifiedName~PpjV1CompilesCanonicalPresentationProgramDeterministically
```

Passed 1/1. The existing comprehensive authored test now compiles a roots-only
treemap and a two-level sunburst, proves omitted descendants are absent from the
native groups and snapshot-free projection, and retains exact authored PPJ
through its existing byte-level program recovery assertion. The test build
emitted one pre-existing nullable warning in the same large test file.

```text
node skills/presentations/skills/presentation-skill-maintainer/scripts/maintain-presentation-skill.mjs check
npx openspec validate ppj-hierarchy-level-controls --strict
node -e "JSON.parse(...schema...); JSON.parse(...registry...)"
git diff --check
```

All passed. The maintainer reported 151 Help APIs, 73 native leaves and 13
host-only operations with no orphaned capability documentation.

## Deliberately unverified here

No full npm suite, package/release gate, Windows playback, hierarchy matrix or
six-sample import runner was executed for this bounded authored-language slice.
