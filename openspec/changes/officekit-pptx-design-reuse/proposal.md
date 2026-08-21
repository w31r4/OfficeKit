## Why

Lossless import is only useful to an Agent when the imported deck can become a
source of design decisions and reusable material. The current benchmark proves
bounded edits and source-package preservation, but it does not yet expose a
stable, evidence-backed description of a deck's visual language or distinguish
a reusable component from an opaque object.

## What Changes

- Add a deterministic PPTX design-profile artifact built from source-package
  evidence plus the bounded semantic projection.
- Record canvas, palette occurrences, typeface and size evidence, density,
  layout families, slide archetypes, recurring geometry candidates, and opaque
  native-object summaries without copying the source deck into the repository.
- Keep the profile descriptive and source-bound: it may recommend a layout or
  component, but it does not grant edit capability or flatten unknown OOXML.
- Define a follow-up source-derived reuse operation separately from profiling;
  the first implementation may reuse only a codec-proven slide/component graph.

## Capabilities

### New Capabilities

- `pptx-design-profile`: Stable design-language evidence and recurring layout /
  component candidates for an imported PPTX.
- `pptx-source-derived-reuse`: Planned, bounded reuse of an imported slide or
  component graph with ownership, source-hash, and review preconditions.

### Modified Capabilities

- None in the public Office wire protocol. The first profile implementation is
  a deterministic repository evaluator and does not change the PPTX codec.

## Impact

- Adds a bounded profile generator and benchmark evidence for the three real
  PPTX samples.
- Does not add a universal AST, HTML conversion path, raw OOXML editor, or
  automatic claim that a repeated geometry is safe to clone.
- Windows PowerPoint and black-box Agent acceptance remain separate gates.
