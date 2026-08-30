# Verification evidence

The existing comprehensive PPJ contract now projects three imported pages with
stable SlidePart-derived IDs, performs one complete three-page permutation and
a matching two-section repartition, then projects the output again. The build
reports only `ppt/presentation.xml`; every page ID and unchanged page-local
element ID survives, the legacy comment stays attached to its page, and the
custom-show page list remains unchanged.

The first focused run exposed the PPJ route-set predicate and was corrected.
The second exposed a real native transaction boundary: reading sections after
moving `p:sldId` temporarily made the old partition opaque. The writer now
updates and verifies the modeled section graph against stable requested page
IDs before moving the exact slide-ID records. The unchanged comprehensive
contract then passed.

```text
dotnet test native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj \
  --no-restore \
  --filter FullyQualifiedName~PpjV1CompilesCanonicalPresentationProgramDeterministically

Passed: 1, Failed: 0

dotnet build native/OfficeKit/src/OfficeKit.Codec/OfficeKit.Codec.csproj --no-restore
Build succeeded. 0 Warning(s), 0 Error(s)

node skills/presentations/skills/presentation-skill-maintainer/scripts/maintain-presentation-skill.mjs check
Presentation Skill maintenance check ok · 151 Help APIs · 73 native leaves · 13 host-only operations

openspec validate ppj-slide-reorder-parity --strict
Change 'ppj-slide-reorder-parity' is valid
```

No new test file, fixture, protobuf field, wire version, full suite, package
gate, or cross-platform matrix was added or run. The one existing comprehensive
test is already heavy and is not promoted to a new routine gate by this change.

The four implementation/evidence commits were pushed normally and remote
`main` fast-forwarded to `9cb171fb040047ef853bb7334cd668892d5f8c5f`
before the final checklist closure. The delegated preset-geometry head
`92459162a97da8eb9af40e6ce1850411cad9f02e` remains an ancestor.
