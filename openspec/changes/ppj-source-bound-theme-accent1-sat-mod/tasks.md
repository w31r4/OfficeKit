## 1. OpenSpec and contract

- [x] 1.1 Add the source-bound accent1 satMod proposal, delta spec, and design, then verify `openspec validate ppj-source-bound-theme-accent1-sat-mod --strict` passes
- [x] 1.2 Add `setThemeAccent1SatMod` and `accentTransforms.accent1.satMod` to the PPJ schema, semantic capability map, registry, and generated/reference documentation; verify schema and capability coverage checks pass

## 2. Native implementation

- [x] 2.1 Read a strict direct accent1 `a:satMod` leaf during PPTX projection and expose the authored thousandth value; verify focused projection assertions pass
- [x] 2.2 Compile one authorized satMod field by patching only the existing direct `a:satMod/@val`, rejecting unsupported topology and combined edits; verify the native codec builds
- [x] 2.3 Add the source-bound satMod regression covering projection, no-op bytes, one-token edit, re-projection, deletion, siblings, combined edits, range, and capability tampering; verify the focused test passes

## 3. Evidence and publication

- [x] 3.1 Update backlog, coverage, capability matrix, and presentation Skill evidence for the new bounded field; verify generated matrix, maintainer, and reference-sync checks pass
- [x] 3.2 Run the focused and theme source-bound tests plus protocol/OpenSpec gates, record any environment-dependent reference-skill skip, and verify the atomic commit contains only this change

Validation notes: codec and test builds succeeded with the repository's existing nullable/analyzer warnings; the focused test passed 1/1 and the complete `PpjSourceBoundTheme` filter passed 26/26. Preview capability coverage, reference Skill sync, maintainer check, capability matrix check, protocol lint/generation, Claude plugin smoke, and strict OpenSpec validation passed. `node test/reference-skills.mjs` remains environment-blocked because `pdftoppm` is not installed (`spawn pdftoppm ENOENT`).
