# Verification evidence

The existing comprehensive PPJ contract now creates two sections and one
custom show, projects item-level capabilities, then performs one source-bound
route transaction. It renames and repartitions the sections, renames and
reorders the custom show with a repeated page, reports only
`ppt/presentation.xml`, and recovers the exact route state after a second
projection.

```text
dotnet test native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj \
  --no-restore \
  --filter FullyQualifiedName~PpjV1CompilesCanonicalPresentationProgramDeterministically

Passed: 1, Failed: 0

dotnet build native/OfficeKit/src/OfficeKit.Codec/OfficeKit.Codec.csproj --no-restore
Build succeeded. 0 Warning(s), 0 Error(s)

node skills/presentations/skills/presentation-skill-maintainer/scripts/maintain-presentation-skill.mjs check
Presentation Skill maintenance check ok · 151 Help APIs · 73 native leaves · 13 host-only operations

openspec validate ppj-section-custom-show-parity --strict
Change 'ppj-section-custom-show-parity' is valid
```

No new test file, fixture, protobuf field, wire version, full suite, package
gate, or cross-platform matrix was added or run.
