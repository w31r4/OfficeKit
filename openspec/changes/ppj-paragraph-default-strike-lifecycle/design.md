## Context

See proposal.md. Authored lowering maps boolean strike values to sngStrike/noStrike; default-run readers and projection retain canonical noStrike/sngStrike/dblStrike strings. The existing default-style cleanup already removes only modeled strike.

## Goals / Non-Goals

Complete direct paragraph strike presence for ordinary text/shape owners without rewriting text. Direct run overrides, table defaults, inherited placeholders and host glyph rendering retain their separate profiles.

## Decisions

Extend exact field masks/authority and the cloned default style. Reuse authored alias lowering and native enum validation. Patch only changed strike in the scalar writer instead of rebuilding unrelated font/effect XML; assign null directly for deletion. Explicit noStrike stays present and differs from omission.

Reject replacing unmodeled source strike. Preserve it during no-op/unrelated scalar assignment/removal, relying on existing source-warning preservation rather than discarding unknown enum values.

## Risks / Trade-offs

Boolean aliases may differ from fresh projection → assert canonical string output. Cancellation may collapse into omission → verify native/fresh PPJ presence. Unknown native strike may disappear during unrelated deletion → one source fixture compares all non-target XML/ZIP and original bytes. Host typography remains unverified.
