## 1. OpenSpec and contract

- [x] 1.1 Add the source-bound accent2 tint proposal, delta spec, and design, then verify `openspec validate ppj-source-bound-theme-accent2-tint --strict` passes
- [x] 1.2 Add `setThemeAccent2Tint` and `accentTransforms.accent2.tint` to the PPJ schema, semantic capability map, registry, and generated/reference documentation; verify schema and capability coverage checks pass

## 2. Native implementation

- [x] 2.1 Read a strict direct accent2 `a:tint` leaf during PPTX projection and expose the authored fraction; verify focused projection assertions pass
- [x] 2.2 Compile one authorized accent2 tint field by patching only the existing direct `a:tint/@val`, rejecting unsupported topology and combined edits; verify the native codec builds
- [x] 2.3 Add the source-bound accent2 tint regression covering projection, no-op bytes, one-token edit, re-projection, deletion, siblings, combined edits, range, and capability tampering; verify the focused test passes

## 3. Evidence and publication

- [x] 3.1 Update backlog, coverage, capability matrix, and presentation Skill evidence for the new bounded field; verify generated matrix, maintainer, and reference-sync checks pass
- [x] 3.2 Run the focused and theme source-bound tests plus protocol/OpenSpec gates, record any environment-dependent reference-skill skip, and verify the atomic commit contains only this change

Validation: codec and test projects build with 0 errors; `PpjSourceBoundThemeAccent2Tint` passes 1/1 and the source-bound theme group passes 36/36. PPJ capability coverage, reference-sync, portability, generated capability matrix, presentation Skill maintenance, icon, PNG, protocol, and Claude marketplace checks pass. `node test/reference-skills.mjs` remains environment-limited because `pdftoppm` is unavailable (`spawn pdftoppm ENOENT`).
