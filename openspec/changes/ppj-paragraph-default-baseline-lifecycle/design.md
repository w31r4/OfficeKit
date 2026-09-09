## Context

See proposal.md. Authored/default-run readers already use FontBaselinePercent within -400..400. The native attribute stores thousandths of a percent. Structured source-bound paragraph edits currently lack its exact field authority.

## Goals / Non-Goals

Complete direct baseline presence for ordinary text/shape paragraphs. Run baseline overrides, table defaults, inherited placeholders and host layout retain separate profiles.

## Decisions

Extend the existing exact field masks/authority and cloned default style. Validate range before converting through native integer thousandths with ties-to-even rounding; tiny negative values become explicit canonical zero. This aligns requested semantics with reprojection.

Patch only changed baseline in the scalar writer, keeping unrelated style/effect XML. Use direct null assignment for removal. Reject replacement of unmodeled source baseline; default-style cleanup clears only modeled baseline so unrelated field deletion preserves source-owned values.

## Risks / Trade-offs

Percent may be confused with points → document units and assert native thousandths. Zero may collapse into omission → verify native/fresh PPJ presence. Negative precision may drift → cover ties and near-zero values. Unknown source values may be lost → cover no-op, replacement rejection and unrelated assignment/removal. Host typography remains unverified.
