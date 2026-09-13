## 1. OpenSpec and contract

- [x] 1.1 Add the source-bound accent1 hueMod proposal, delta spec, and design, then verify `openspec validate ppj-source-bound-theme-accent1-hue-mod --strict` passes
- [x] 1.2 Add `setThemeAccent1HueMod` and `accentTransforms.accent1.hueMod` to the PPJ schema, semantic capability map, registry, and generated/reference documentation; verify schema and capability coverage checks pass

## 2. Native implementation

- [x] 2.1 Read a strict direct accent1 `a:hueMod` leaf during PPTX projection and expose the authored thousandth value; verify focused projection assertions pass
- [x] 2.2 Compile one authorized hueMod field by patching only the existing direct `a:hueMod/@val`, rejecting unsupported topology and combined edits; verify the native codec builds
- [x] 2.3 Add the source-bound hueMod regression covering projection, no-op bytes, one-token edit, re-projection, deletion, siblings, combined edits, range, and capability tampering; verify the focused test passes

## 3. Evidence and publication

- [x] 3.1 Update backlog, coverage, capability matrix, and presentation Skill evidence for the new bounded field; verify generated matrix, maintainer, and reference-sync checks pass
- [x] 3.2 Run the focused and theme source-bound tests plus protocol/OpenSpec gates, record any environment-dependent reference-skill skip, and verify the atomic commit contains only this change

Validation: codec and test projects build with 0 errors; `PpjSourceBoundThemeAccent1HueMod` passes 1/1 and the source-bound theme group passes 34/34. PPJ capability coverage, reference-sync, portability, generated capability matrix, presentation Skill maintenance, icon, PNG, protocol, and Claude marketplace checks pass. `node test/reference-skills.mjs` remains environment-limited because `pdftoppm` is unavailable (`spawn pdftoppm ENOENT`).
