## 1. PPJ and codec contract

- [x] 1.1 Add `textShadowColorScheme` to native projection, source-bound
  validation/read/patch dispatch, and the PPJ schema/registry/generated matrix;
  verify a canonical direct-run scheme token projects as a typed leaf.
- [x] 1.2 Reuse the strict direct-run shadow owner proof and splice only
  `outerShdw/schemeClr/@val`; verify RGB, transformed, malformed, and compound
  graphs remain outside this leaf.

## 2. Regression evidence

- [x] 2.1 Add a focused authored → theme-token source fixture → source-bound edit
  → second projection test with SlidePart-only and non-target ZIP preservation.
- [x] 2.2 Cover stale, unchanged, malformed, RGB, transformed, missing-geometry,
  sibling-effect, unknown-descendant, and color-transform boundaries.

## 3. Evidence and gates

- [x] 3.1 Update the presentation text Skill, coverage, and P0/P1 backlog notes;
  verify generated docs and OpenSpec strict validation.
- [x] 3.2 Run the narrow codec build/test, schema/matrix, portability,
  reference-sync, preview capability/field-limit, and presentation-maintainer
  gates before the atomic commit.
