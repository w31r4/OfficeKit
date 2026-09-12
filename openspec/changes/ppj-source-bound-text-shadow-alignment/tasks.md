## 1. PPJ surface

- [x] 1.1 Add `textShadowAlignment` to the strict direct-run projection and source-bound validation/read/patch dispatch; verify the existing `algn` token is re-proven before replacement.
- [x] 1.2 Add the bounded string leaf to the PPJ schema, capability registry, generated capability matrix, and presentation Skill/reference coverage; verify generated checks pass.

## 2. Regression evidence

- [x] 2.1 Add a focused direct-run authored/source-bound/second-projection fixture for RGB and theme-color shadows; verify only the owning SlidePart and `outerShdw/@algn` change.
- [x] 2.2 Add missing/invalid alignment, missing geometry, transform, sibling-effect, unknown-descendant, stale-proof, and no-op cases; verify omission or fail-closed behavior.

## 3. OpenSpec and gates

- [x] 3.1 Record the completed behavior in backlog and coverage notes; validate with `openspec validate ppj-source-bound-text-shadow-alignment --strict`.
- [x] 3.2 Run the narrow affected codec, portability, reference, preview, and generated-document gates and inspect the staged diff before the atomic commit.
