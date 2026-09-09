## Context

See proposal.md. Authored/default-run readers already use FontSpacingPoints with a finite -768..768pt range. The native spc attribute stores hundredths; structured source-bound paragraph edits currently lack its exact field authority.

## Goals / Non-Goals

Complete the direct default field lifecycle for ordinary text/shape paragraphs. Run spacing, table defaults, inherited placeholders and host text layout retain their separate profiles.

## Decisions

Extend the existing exact field masks/authority and cloned default style. Validate the requested range before rounding to native hundredths with ties to even, including negative values. Convert through the native integer representation so a small negative value that rounds to -0 becomes explicit canonical zero. This keeps requested semantics aligned with reprojection.

Patch only changed spacing in the existing scalar writer instead of rebuilding unrelated style/effects. Use direct null assignment for removal. Reject replacing an unmodeled native spc; default-style cleanup removes only modeled spacing, preserving source-owned values during unrelated field deletion.

## Risks / Trade-offs

Zero may collapse into omission → assert native and fresh PPJ presence. Signed precision may drift → cover positive/negative ties and sub-hundredth values. Unknown source values may be lost during cleanup → cover no-op, replacement rejection, unrelated assignment/removal. Preview typography remains partial.
