# Evidence

## Implemented contract

- A `combo` containing scatter or bubble data compiles as one bounded editable
  DrawingML group on a shared value/value coordinate system. Ordinary
  categorical combo charts continue to use native ChartParts.
- A candlestick accepts up to four aligned line, area or column overlays.
  Filled overlays render behind the price marks and line overlays render above
  them.
- Stable generated element IDs preserve editable marks. Embedded PPJ restores
  the exact authored semantics; snapshot-free import reports the ordinary group
  rather than guessing chart data.
- Unsupported axes, fields, invisible scatter marks, ignored marker settings,
  formulas and chart-build animation fail before output is written.

## Narrow verification

All commands ran from the isolated feature worktree.

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

Passed 1/1. The existing comprehensive authored-PPJ test now proves numeric
combo validation, native editable z-order, candlestick moving-average overlay,
exact embedded recovery, deterministic output and honest snapshot-free group
projection. The test build emitted one pre-existing nullable warning in the
same large test file.

```text
node skills/presentations/skills/presentation-skill-maintainer/scripts/maintain-presentation-skill.mjs check
npx openspec validate ppj-mixed-cartesian-overlays --strict
node -e "JSON.parse(...schema...); JSON.parse(...registry...)"
git diff --check
```

All passed. The maintainer reported 151 Help APIs, 73 native leaves and 13
host-only operations with no orphaned capability documentation.

## Deliberately unverified here

No full npm suite, package/release gate, Windows PowerPoint playback, broad
chart matrix or six-sample PPTX runner was executed for this bounded slice.
Those remain final-goal or release evidence rather than daily feature gates.
