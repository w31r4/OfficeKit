# Verification evidence

The existing PPJ state-edit transaction now renames and hides its source page
while changing one element's hidden/locked state. It still reports only
`ppt/slides/slide1.xml`, includes both page and element stable IDs in the
footprint, and recovers the exact page name, page visibility, and element state
after second projection.

```text
dotnet test native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj \
  --no-restore \
  --filter FullyQualifiedName~PpjV1CompilesCanonicalPresentationProgramDeterministically

Passed: 1, Failed: 0

dotnet build native/OfficeKit/src/OfficeKit.Codec/OfficeKit.Codec.csproj --no-restore
Build succeeded. 0 Warning(s), 0 Error(s)

node skills/presentations/skills/presentation-skill-maintainer/scripts/maintain-presentation-skill.mjs check
Presentation Skill maintenance check ok · 151 Help APIs · 73 native leaves · 13 host-only operations

npx openspec validate ppj-page-metadata-parity --strict
Change 'ppj-page-metadata-parity' is valid
```

No new test file, fixture, protobuf field, wire version, full suite, package
gate, or cross-platform matrix was added or run.
