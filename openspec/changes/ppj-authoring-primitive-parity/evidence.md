# Evidence

## Focused contract

Command:

```text
dotnet test native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj \
  --filter FullyQualifiedName~PpjV1CompilesCanonicalPresentationProgramDeterministically \
  -p:UseSharedCompilation=false --no-restore --blame-hang-timeout 5m
```

Result: one contract passed. It compiled and inspected table inheritance,
Sankey right alignment, a named node-color override, and existing deterministic
reimport behavior. The same contract rejected an undeclared Sankey color-map
key. The command was intentionally not expanded into a full suite.

## Language and Agent surface

- Generated PPJ reference synchronized from the schema: 2,233 lines, 732
  documented root/definition fields, 178 preset names.
- Presentation Skill maintainer check passed: 151 Help APIs, 73 native leaves,
  13 host-only operations.
- Strict OpenSpec validation passed.
- Connector remains the sole ordinary line primitive; no duplicate PPJ element
  type or wire operation was introduced.

## Boundaries

This no-wire batch does not add named icons, formula representation, remote
asset fetching, arbitrary table selectors, or the broader native chart-writer
work listed in the audit. No full `npm test`, package gate, hosted CI, or host
playback acceptance was run.
