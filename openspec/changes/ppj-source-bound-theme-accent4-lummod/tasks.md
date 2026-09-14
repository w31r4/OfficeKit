## 1. OpenSpec and contract

- [x] 1.1 Add the source-bound accent4 lumMod proposal, delta spec, and design, then verify `openspec validate ppj-source-bound-theme-accent4-lummod --strict` passes
- [x] 1.2 Add `setThemeAccent4LumMod` and `accentTransforms.accent4.lumMod` to the PPJ schema, semantic capability map, registry, and generated/reference documentation; verify schema and capability coverage checks pass

## 2. Native implementation

- [x] 2.1 Read a strict direct accent4 `a:lumMod` leaf during PPTX projection and expose the authored fraction; verify focused projection assertions pass
- [x] 2.2 Compile one authorized accent4 lumMod field by patching only the existing direct `a:lumMod/@val`, rejecting unsupported topology and combined edits; verify the native codec builds
- [x] 2.3 Add the source-bound accent4 lumMod regression covering projection, no-op bytes, one-token edit, re-projection, deletion, siblings, combined edits, range, and capability tampering; verify the focused test passes

## 3. Evidence and publication

- [x] 3.1 Update backlog, coverage, capability matrix, and presentation Skill evidence for the new bounded field; verify generated matrix, maintainer, and reference-sync checks pass
- [x] 3.2 Run the focused and theme source-bound tests plus protocol/OpenSpec gates, record any environment-dependent reference-skill skip, and verify the atomic commit contains only this change

Validation: codec and test projects build with 0 errors; `PpjSourceBoundThemeAccent4LumMod` passes 1/1 and the source-bound theme group passes 70/70. PPJ capability coverage, reference-sync, portability, generated capability matrix, presentation Skill maintenance, icon, PNG, Buf lint, proto generation, and Claude marketplace checks pass. `npm run proto:check` cannot run from the archive candidate because it requires a Git worktree for its final generated-file diff; `buf lint` and `proto:generate` are run separately. `node test/reference-skills.mjs` remains environment-limited because `pdftoppm` is unavailable (`spawn pdftoppm ENOENT`).
