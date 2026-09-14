## 1. OpenSpec and contract

- [x] 1.1 Add the source-bound accent4 hueMod proposal, delta spec, and design, then verify `openspec validate ppj-source-bound-theme-accent4-huemod --strict` passes
- [x] 1.2 Add `setThemeAccent4HueMod` and `accentTransforms.accent4.hueMod` to the PPJ schema, semantic capability map, registry, and generated/reference documentation; verify schema and capability coverage checks pass

## 2. Native implementation

- [x] 2.1 Read a strict direct accent4 `a:hueMod` leaf during PPTX projection and expose the fraction; verify focused projection assertions pass
- [x] 2.2 Compile one authorized accent4 hueMod field by patching only the existing direct `a:hueMod/@val`, rejecting unsupported topology and combined edits; verify the native codec builds
- [x] 2.3 Add the source-bound accent4 hueMod regression covering projection, no-op bytes, one-token edit, re-projection, deletion, siblings, combined edits, range, and capability tampering; verify the focused test passes

## 3. Evidence and publication

- [x] 3.1 Update backlog, coverage, capability matrix, and presentation Skill evidence for the new bounded field; verify generated matrix, maintainer, and reference-sync checks pass
- [x] 3.2 Run the focused and theme source-bound tests plus protocol/OpenSpec gates, record any environment-dependent reference-skill skip, and verify the atomic commit contains only this change

Validation: codec and test projects build successfully with 0 errors (the
repository's existing analyzer warnings remain); the focused
`PpjSourceBoundThemeAccent4HueMod` test passes 1/1 and the source-bound theme
group passes 82/82. PPJ capability coverage, reference-sync, portability,
generated capability matrix, presentation Skill maintenance, icon, PNG, Buf
lint, proto generation, and Claude marketplace checks pass. `npm run
proto:check` cannot complete from the archive candidate because its final
generated-file diff requires a Git worktree; `buf lint` and `proto:generate`
are run separately. `node test/reference-skills.mjs` remains
environment-limited because `pdftoppm` is unavailable (`spawn pdftoppm
ENOENT`).
