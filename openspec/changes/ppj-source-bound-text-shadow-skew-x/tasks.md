## 1. PPJ surface

- [x] 1.1 Add `textShadowSkewX` to direct-run parsing, projection, safety proof, readback, dispatch, token validation, and `kx` patching.
- [x] 1.2 Add the text-run-specific schema field, capability registry entry, generated matrix, and presentation Skill/reference coverage.

## 2. Regression evidence

- [x] 2.1 Add an authored → imported/source-bound → second-projection fixture covering negative and positive signed horizontal skew values and XML/package preservation.
- [x] 2.2 Add missing/invalid token, other-transform, sibling/unknown-child, stale-proof, and no-op cases.

## 3. OpenSpec and gates

- [x] 3.1 Record the increment in backlog and coverage notes; validate with `openspec validate ppj-source-bound-text-shadow-skew-x --strict`.
- [x] 3.2 Run the narrow codec, portability, reference, preview, and generated-document gates and inspect the staged diff before atomic publication.
