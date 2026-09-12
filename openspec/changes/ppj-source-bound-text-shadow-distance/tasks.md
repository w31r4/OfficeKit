## 1. PPJ surface

- [x] 1.1 Add `textShadowDistanceEmu` to native projection and source-bound validation/read/patch dispatch; verify canonical distance values are projected and only `outerShdw/@dist` is replaced.
- [x] 1.2 Add the bounded field to the PPJ schema, capability registry, generated capability matrix, and presentation Skill/reference coverage; verify JSON/schema and generated-doc checks pass.

## 2. Regression evidence

- [x] 2.1 Add a focused authored → source-bound edit → second projection fixture for direct rich-text shadow distance, including SlidePart-only and non-target ZIP preservation; verify the focused codec test passes.
- [x] 2.2 Add malformed, missing, transformed, sibling-effect, overflow, and stale-token cases; verify they remain opaque or fail closed and Open XML validation remains clean.

## 3. OpenSpec and gates

- [x] 3.1 Record the completed behavior in the backlog and coverage notes and validate this change with `openspec validate ppj-source-bound-text-shadow-distance --strict`.
- [x] 3.2 Run the narrow affected portability/reference/preview gates and inspect the staged diff before the atomic commit.
