## 1. Capability and projection

- [x] 1.1 Add `setAnimations` to the closed PPJ native capability vocabulary.
- [x] 1.2 Issue it only for canonical editable/addable non-Morph timing state.

## 2. Source-bound lowering

- [x] 2.1 Reuse the authored animation lowerer for the requested array.
- [x] 2.2 Map stable PPJ targets to exact top-level or group-descendant wire IDs.
- [x] 2.3 Preserve opaque timing and reject stale, Morph, or overlay-mixed edits.

## 3. Agent surface and lean verification

- [x] 3.1 Regenerate `ppj.md` and update motion/import guidance and coverage.
- [x] 3.2 Extend one existing source PPJ round-trip with animation add/reimport.
- [x] 3.3 Run the focused test, Skill maintainer, and strict OpenSpec check once.
- [x] 3.4 Commit atomically and fast-forward main without force pushing.

## Evidence

- `PpjSourceBoundProgramReusesOneProvenSlide` passed after adding one paragraph
  wipe to a source rich-text title. The source had no timing; PPJ discovered
  `setAnimations/animations`, build changed only that SlidePart and wrote
  native `p:timing`, and second projection recovered the same target, wipe,
  and paragraph build plus a fresh edit capability.
- The first focused run exposed a missing closed capability-field enum entry.
  The next run exposed that the timing postwrite hash incorrectly included an
  Agent animation ID and derived target-kind label that native OOXML does not
  round-trip. Both stable contract defects were fixed before the passing run;
  the postwrite oracle remains strict over actual playback semantics.
- `presentation-skill-maintainer check` passed with 151 Help APIs, 73 native
  leaves, and 13 host-only operations after regenerating `ppj.md`.
- `openspec validate ppj-source-motion-parity --strict` passed.
- No new test file, timing writer, Office wire field/version, procedural motion
  DSL, playback claim, full suite, package gate, or sample matrix was added.
