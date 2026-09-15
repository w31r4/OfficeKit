## 1. OpenSpec and contract

- [x] 1.1 Add the source-bound accent1 gray proposal and design, then verify `openspec validate ppj-source-bound-theme-accent1-gray --strict` passes
- [x] 1.2 Add `setThemeAccent1Gray` and `accentTransforms.accent1.gray` to the PPJ schema, semantic capability map, registry, and generated/reference documentation; verify schema and capability coverage checks pass

## 2. Native implementation

- [x] 2.1 Read a strict direct accent1 `a:gray` leaf during PPTX projection and expose the boolean field; verify focused projection assertions pass
- [x] 2.2 Compile one authorized gray field by removing only the existing direct `a:gray` child, rejecting absent owners and combined edits; verify the native codec builds
- [x] 2.3 Add the source-bound gray regression covering projection, no-op bytes, removal, re-projection, deletion, siblings, invalid types, combined edits, and capability tampering; verify the focused test passes

## 3. Evidence and publication

- [x] 3.1 Update backlog, coverage, capability matrix, and presentation Skill evidence for the new bounded field; verify generated matrix, maintainer, and reference-sync checks pass
- [x] 3.2 Run the focused and theme source-bound tests plus protocol/OpenSpec gates, record any environment-dependent reference-skill skip, and verify the atomic commit contains only this change

Validation: codec and test projects build with 0 errors; the focused source-bound gray test passes 1/1 and the source-bound theme group passes 129/129; PPJ capability coverage, reference-sync, portability, generated capability matrix, presentation Skill maintenance, and OpenSpec checks pass. The standalone `buf lint` gate is unavailable in this environment because `buf` is not installed. Full host rendering remains outside this bounded experiment.
