## 1. Projection

- [x] 1.1 Project a proven diagram binding as source-bound SmartArt.
- [x] 1.2 Preserve node/run boundaries and issue bounded nativeRef authority.

## 2. Source-bound lowering

- [x] 2.1 Validate unchanged node identity and run topology.
- [x] 2.2 Lower only changed text into the existing diagram binding.
- [x] 2.3 Keep unsupported diagrams opaque and fail closed on stale authority.

## 3. Agent surface and lean verification

- [x] 3.1 Regenerate `ppj.md` and update imported-edit guidance and coverage.
- [x] 3.2 Extend one existing SmartArt test with PPJ projection/edit/reimport.
- [x] 3.3 Run the focused test, Skill maintainer, and strict OpenSpec check once.
- [x] 3.4 Commit atomically and fast-forward main without force pushing.

## Evidence

- `SourceBoundSmartArtPlainNodeTextCanBeEditedWithoutChangingItsGraph` passed
  after projecting the existing two-node source diagram as typed PPJ, changing
  one `nodes[].text` string, compiling, and projecting the result again.
- The PPJ build reported only `ppt/diagrams/clone-data.xml` as changed. The
  SlidePart, slide relationships, layout, quick-style, and colors parts stayed
  byte-identical; second projection recovered `PPJ revised node` and a fresh
  SmartArt edit capability.
- The same native test already proves unsafe model-ID and node-count changes,
  connected diagram graphs, malformed IDs, and annotated text remain rejected
  by the underlying source-bound codec.
- `presentation-skill-maintainer check` passed with 151 Help APIs, 73 native
  leaves, and 13 host-only operations after regenerating `ppj.md`.
- `openspec validate ppj-source-smartart-parity --strict` passed.
- No new schema field, wire field/version, DiagramML writer, raw XML surface,
  test file, sample matrix, full suite, package gate, or playback claim was
  added.
