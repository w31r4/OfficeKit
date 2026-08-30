## 1. Language and projection

- [x] 1.1 Add the finite `sourceClone` page descriptor and validation boundary.
- [x] 1.2 Issue `duplicate/pageClone` only from a proven slide clone capability.

## 2. Source-bound lowering

- [x] 2.1 Resolve the exact source page and capability from the fresh projection.
- [x] 2.2 Lower one immutable adjacent clone through the existing native writer.
- [x] 2.3 Keep section, custom-show, comment, and source-page state unchanged.

## 3. Agent surface and lean verification

- [x] 3.1 Regenerate `ppj.md` and update continuation guidance and coverage.
- [x] 3.2 Add one focused contract in the existing PPJ codec test file.
- [x] 3.3 Run focused C# test, Skill-maintainer, and strict OpenSpec checks.
- [ ] 3.4 Commit atomically and fast-forward main without force pushing.

## Evidence

- `PpjSourceBoundProgramReusesOneProvenSlide` passed once. It authored one
  minimal page, removed the embedded OfficeKit program to create a third-party
  source, projected `duplicate/pageClone`, built one adjacent clone through
  NativeAOT, proved the original SlidePart remained in place, proved source
  and clone Slide XML were equal, and reprojected both pages as ordinary
  non-macro source-bound content.
- `presentation-skill-maintainer check` passed with 151 Help APIs, 73 native
  leaves, and 13 host-only operations after regenerating `ppj.md`.
- `openspec validate ppj-source-slide-reuse-parity --strict` passed.
- No new test file, clone writer, wire field, full suite, package gate, sample
  matrix, raw OOXML surface, or procedural PPJ operation list was added.
