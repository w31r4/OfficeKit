## 1. Source evidence and route

- [x] 1.1 Freeze the three external PPTX hashes and source/profile evidence
  without adding the source binaries to the repository.
- [x] 1.2 Add the public, bounded template-conditioned generation guidance.
- [x] 1.3 Add a source-free fixture smoke using only the public OfficeKit API.

## 2. Generation and preservation

- [x] 2.1 Select clone-safe source archetypes and bounded text/SVG targets.
- [x] 2.2 Cross export/reimport boundaries when reusing a source slide and use
  source ordinal/occurrence locators after public ids are regenerated.
- [x] 2.3 Generate ten pages per real sample, reimport the result, and verify
  target values and source protection.
- [x] 2.4 Compare non-target package parts and logical source slide parts; record
  topology additions separately from retained-part drift.
- [x] 2.5 Compare output review issues with the source baseline and report
  renderer limitations explicitly.

## 3. Independent acceptance

- [x] 3.1 Run three fresh black-box Agent tasks using only the portable Skill,
  public package, and each external source; store frame map, profile, plan,
  output, reimport, and review evidence in
  `evals/pptx-generation/agent-blackbox.v1.json`. The portable lane records
  structural/layout review separately from unavailable visual review.
- [x] 3.2 Add a conversational local edit after generation and verify that the
  generated output can be reopened and safely changed without source drift.
- [x] 3.3 Run the packed clean-install route for all three sources and record
  available renderer evidence. The blue-gray sample has no portable renderer
  because its custom geometry exceeds the renderer's bounded path profile;
  this remains `visualReview: "unavailable"`, not a visual pass.

## 4. Release gates

- [x] 4.1 Add the generation evidence smoke to the Presentation slow-gate
  segment after the benchmark evidence format is stable.
- [x] 4.2 Update coverage, OpenSpec evidence, release notes, and the slow-gate
  inventory. No public API or Help surface changed in this evidence-only
  slice, so generated API/help files remain unchanged; package contents are
  checked by the existing pack gate.
- [x] 4.3 Run full npm, package, deterministic WASM, and hosted CI gates, then
  deliver one atomic ordinary push. Local fast (38/38), Presentation slow
  (8/8), release slow (5/5), proto/WASM reproducibility, and hosted slow run
  32567677699 are green; the evidence commit is pushed on the isolated branch.
