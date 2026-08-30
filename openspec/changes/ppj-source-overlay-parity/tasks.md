## 1. Language and projection

- [x] 1.1 Add `appendElement/elements` to the PPJ capability vocabulary.
- [x] 1.2 Issue the capability on ordinary source-bound pages.

## 2. Source-bound lowering

- [x] 2.1 Recognize an unchanged source prefix plus a fresh typed suffix.
- [x] 2.2 Lower bounded text, shape, and image overlays through existing code.
- [x] 2.3 Reject interleaving, unsupported types, and mixed slide mutations.

## 3. Agent surface and lean verification

- [x] 3.1 Regenerate `ppj.md` and update import guidance and coverage.
- [x] 3.2 Add one focused source overlay contract to the existing PPJ test file.
- [x] 3.3 Run the focused test, Skill maintainer, and strict OpenSpec check once.
- [x] 3.4 Commit atomically and fast-forward main without force pushing.

## Evidence

- `PpjSourceBoundProgramReusesOneProvenSlide` passed once after its existing
  slide-clone proof was extended through the required build/reimport boundary.
  The reprojected clone issued `appendElement/elements`; PPJ appended one
  editable textbox; only that clone's actual SlidePart changed; the unrelated
  page XML stayed exact; and a second projection returned the text as a typed
  element with nativeRef.
- The compiler initially exposed two useful contract mistakes during this one
  run: PPJ overlay elements must normalize absent hidden/locked state to native
  `false`, and OPC SlidePart filenames are not presentation page numbers. Both
  were corrected in the implementation/test rather than hidden by a fixture.
- `presentation-skill-maintainer check` passed with 151 Help APIs, 73 native
  leaves, and 13 host-only operations after regenerating `ppj.md`.
- `openspec validate ppj-source-overlay-parity --strict` passed.
- No new test file, overlay writer, command DSL, wire field/version, full npm
  suite, package gate, sample matrix, or raw OOXML surface was added.
