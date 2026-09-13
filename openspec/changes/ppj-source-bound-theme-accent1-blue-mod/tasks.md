## 1. OpenSpec and contract

- [x] 1.1 Add the source-bound accent1 blueMod proposal, delta spec, and design, then verify `openspec validate ppj-source-bound-theme-accent1-blue-mod --strict` passes
- [x] 1.2 Add `setThemeAccent1BlueMod` and `accentTransforms.accent1.blueMod` to the PPJ schema, semantic capability map, registry, and generated/reference documentation; verify schema and capability coverage checks pass

## 2. Native implementation

- [x] 2.1 Read a strict direct accent1 `a:blueMod` leaf during PPTX projection and expose the authored thousandth value; verify focused projection assertions pass
- [x] 2.2 Compile one authorized blueMod field by patching only the existing direct `a:blueMod/@val`, rejecting unsupported topology and combined edits; verify the native codec builds
- [x] 2.3 Add the source-bound blueMod regression covering projection, no-op bytes, one-token edit, re-projection, deletion, siblings, combined edits, range, and capability tampering; verify the focused test passes

## 3. Evidence and publication

- [x] 3.1 Update backlog, coverage, capability matrix, and presentation Skill evidence for the new bounded field; verify generated matrix, maintainer, and reference-sync checks pass
- [x] 3.2 Run the focused and theme source-bound tests plus protocol/OpenSpec gates, record any environment-dependent reference-skill skip, and verify the atomic commit contains only this change

## Validation

- `dotnet build native/OfficeKit/src/OfficeKit.Codec/OfficeKit.Codec.csproj -c Release`: passed with 0 errors and the existing nullable warnings.
- `dotnet build native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj -c Release`: passed with 0 errors and the existing test warnings.
- `PpjSourceBoundThemeAccent1BlueMod`: 1 passed; `PpjSourceBoundTheme`: 32 passed.
- PPJ capability coverage, reference Skill sync, Skill portability, generated matrix, presentation Skill maintenance, icon catalog, PNG, protobuf lint/generation, and Claude plugin checks passed.
- `node test/reference-skills.mjs` remains environment-dependent here because `pdftoppm` is unavailable (`spawn pdftoppm ENOENT`).
