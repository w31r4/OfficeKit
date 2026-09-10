## Context

See proposal.md. The native wire already has point, relative and follow-text size cases. PPJ omits follow-text, permits both numeric forms, and exposes 0–1000 bounds that disagree with native validation. The native reader currently uses SDK numeric values directly, which can throw for malformed source tokens.

## Goals / Non-Goals

**Goals:** Reuse the existing size oneof and source paragraph field mutation path. Preserve direct choice presence and native numeric precision.

**Non-Goals:** Resolve inherited text metrics, change marker kinds or expand imported arbitrary marker topology. Host glyph/layout acceptance remains separate.

## Decisions

- Keep the established names and numeric ratio meaning. Align schema bounds with native validation and add an explicit true-only follow-text property; enforce mutual exclusion in schema and lowering.
- Mask all three size properties together, require authority separately for every property actually changed, and copy only the changed paragraph's requested size choice.
- Reuse shared choice apply with modeled deletion enabled. Its existing semantic comparison preserves equivalent source size nodes, including signed/padded numeric spelling.
- Parse only a single direct namespace-free integer `val` before reading numeric size state. Classify bounds and unknown metadata first, preserving malformed/ambiguous source nodes instead of triggering SDK parsing failures.
- Use a small text/shape lifecycle fixture for character/number/picture markers, numeric bounds/precision examples and malformed source sizes. Compare non-target XML/ZIP and retain font/color spelling. Run the related shared list/table/master regressions and explicit preview diagnostics.

## Risks / Trade-offs

- Relative size uses a ratio despite the field name → document 1=100% and test it distinctly from follow-text/absence.
- Shared size deletion affects list/table/master writers → run their focused existing regressions.
- Relative/follow-text preview depends on inherited metrics → report unavailable while retaining native state.
