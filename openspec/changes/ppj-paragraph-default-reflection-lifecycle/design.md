## Context

See proposal.md. The chart reflection schema, builder, projector and native reader already cover all 14 reflection values. Paragraph defaults still use required ordinary reflection geometry and full-span source reading. The preceding inner-shadow change supplies property aliases, direct-effect patching and unchanged-run preservation.

## Goals / Non-Goals

Make all modeled direct paragraph reflection values independently expressible with preserved presence. Keep unsupported source graphs opaque and preserve other owners' existing authoring contracts. Host appearance, inheritance and other default-effect lifecycles remain separate.

## Decisions

Reference chartTextReflection from paragraphDefaultTextStyle and reuse its optional builder/projector. This avoids a second partial reflection vocabulary and keeps existing run/shape/image required geometry unchanged.

Opt the isolated paragraph-default reader into variable positions and transforms. Keep strict native-leaf owner proofs unchanged; only the schema range for an existing default reflection distance leaf needs the 1,270,000,000 EMU endpoint. Source unknown attributes, children, duplicate lists/effects and DAGs remain guarded by direct-effect recognition.

Patch only a changed reflection using the common effect helper. Compare serialized reflection values, including optional presence, because generated protobuf equality equates missing scalar values with zero/false. Empty reflection objects remain real native elements; whole-effect deletion removes only the target and an empty attribute-free list.

## Risks / Trade-offs

Optional numeric or boolean presence collapses → exercise independent removal/restoration from original bytes and inspect native attributes plus fresh PPJ.

Expanded positions/transforms disturb older default-effect tests → retain ordinary full-span behavior and run related reflection/default-style regressions.

Unmodeled siblings or direct runs get rebuilt → preserve their XML and non-target ZIP entries in mixed fixtures.

Preview coverage is mistaken for host validation → keep explicit partial diagnostics and record unperformed host/AOT checks.
