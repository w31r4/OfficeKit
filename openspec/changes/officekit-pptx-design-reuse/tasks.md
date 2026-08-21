## 1. Evidence-backed design profile

- [x] 1.1 Add a deterministic profile generator with source hash, package
  structure, inspected element evidence, palette/typeface/size counts, density,
  layout families, slide archetypes, recurring geometry candidates, and opaque
  native summaries.
- [x] 1.2 Generate and validate profiles for the three real benchmark decks
  without checking the source files into the repository.
- [x] 1.3 Add a smoke test and make the profile evidence part of the PPTX
  benchmark gate.

## 2. Source-derived reuse

- [ ] 2.1 Define a source-bound reuse request with source SHA, selected slide or
  component ID, ownership proof, and expected revision.
- [x] 2.2 Reuse one codec-proven complete slide graph without sharing mutable
  descendants or changing the original source slide.
- [ ] 2.3 Expose recurring component candidates only as inspectable references;
  reject ambiguous, shared, or opaque component graphs before mutation.
- [x] 2.4 Add review, second-import, package-footprint, and native-render checks
  for the bounded reused-slide slice; component candidates remain gated by 2.3.

## 3. Same-style continuation

- [x] 3.1 From each benchmark deck, create one new page using only profiled
  evidence and approved source-derived assets.
- [x] 3.2 Verify that the new page does not modify non-target source parts and
  that the original deck's opaque structures remain intact.
- [ ] 3.3 Run three fresh model Agent tasks through inspect → plan → reuse →
  review → resume, then add Windows PowerPoint evidence before marking this
  change done. A deterministic public-REPL rehearsal covers the protocol slice;
  it is not a substitute for this model/host gate.
