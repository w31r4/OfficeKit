# Verification evidence

The existing comprehensive PPJ transaction now edits one imported chart title
and one bounded legacy comment body. The compiler reports exactly the chart and
legacy-comment parts, records both stable IDs, and a second projection recovers
the new comment text. Comment identity, page binding, author, timestamp,
position, resolution state, order, and nativeRef remain unchanged.

```text
dotnet test native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj \
  --no-restore \
  --filter FullyQualifiedName~PpjV1CompilesCanonicalPresentationProgramDeterministically

Passed: 1, Failed: 0

dotnet build native/OfficeKit/src/OfficeKit.Codec/OfficeKit.Codec.csproj --no-restore
Build succeeded. 0 Warning(s), 0 Error(s)

node skills/presentations/skills/presentation-skill-maintainer/scripts/maintain-presentation-skill.mjs check
Presentation Skill maintenance check ok · 151 Help APIs · 73 native leaves · 13 host-only operations

openspec validate ppj-legacy-comment-text-parity --strict
Change 'ppj-legacy-comment-text-parity' is valid
```

No new test file, fixture, protobuf field, wire version, full suite, package
gate, or cross-platform matrix was added or run.
