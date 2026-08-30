## 1. Language contract

- [x] 1.1 Add optional `sourceClone.retainElement` to schema and typed C# model.
- [x] 1.2 Validate a direct source-page element and every sibling delete proof.

## 2. Source-bound lowering

- [x] 2.1 Retain one exact source element in the clone wire projection.
- [x] 2.2 Lower every sibling through existing element-deletion bindings.
- [x] 2.3 Preserve full-page clone behavior when the selector is absent.

## 3. Agent surface and lean verification

- [x] 3.1 Regenerate `ppj.md` and update continuation guidance and coverage.
- [x] 3.2 Extend one existing clone round-trip with component reuse and reimport.
- [x] 3.3 Run the focused test, Skill maintainer, and strict OpenSpec check once.
- [x] 3.4 Commit atomically and fast-forward main without force pushing.

## Evidence

- `PpjSourceBoundProgramReusesOneProvenSlide` passed once after the existing
  full-page clone contract was extended with an independent component clone.
  It retained one rich-text title, deleted its proven rectangle sibling,
  proved the original SlidePart XML unchanged, proved the clone contained one
  native shape, and reprojected it as one typed text element with nativeRef.
- The same test still builds the legacy two-field `sourceClone` and proves the
  output clone Slide XML equals the source, preserving additive compatibility.
- One narrow `OfficeKit.Codec` build passed before the focused contract.
- `presentation-skill-maintainer check` passed with 151 Help APIs, 73 native
  leaves, and 13 host-only operations after regenerating `ppj.md`.
- `openspec validate ppj-source-component-reuse-parity --strict` passed.
- No new test file, candidate analyzer, native writer, Office wire field,
  procedural command surface, full suite, package gate, or sample matrix was
  added.
