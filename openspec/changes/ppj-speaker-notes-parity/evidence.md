# Verification evidence

## Focused contract

```text
dotnet test native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj \
  --no-restore \
  --filter FullyQualifiedName~PpjV1CompilesCanonicalPresentationProgramDeterministically

Passed: 1, Failed: 0
```

The existing comprehensive PPJ test now proves one source package can:

- project a three-run formatted notes body without flattening;
- issue `setNotes` only on the two native-proven page profiles;
- edit the middle run while preserving bold and RGB style;
- add one plain notes value to a different notes-absent page;
- report both page IDs and notes-related package parts in the footprint; and
- recover both edits after a second PPTX-to-PPJ projection.

## Native and documentation checks

```text
dotnet build native/OfficeKit/src/OfficeKit.Codec/OfficeKit.Codec.csproj --no-restore
Build succeeded. 0 Warning(s), 0 Error(s)

node skills/presentations/skills/presentation-skill-maintainer/scripts/maintain-presentation-skill.mjs check
Presentation Skill maintenance check ok · 151 Help APIs · 73 native leaves · 13 host-only operations

npx openspec validate ppj-speaker-notes-parity --strict
Change 'ppj-speaker-notes-parity' is valid
```

No protobuf field or wire version changed. No full test suite, packaging gate,
or cross-platform matrix was run for this bounded parity slice.
