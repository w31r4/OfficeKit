## Context

See `proposal.md` for the motivation. The current authored compositor accepts
normal opacity only for model types that lower to a shape or image. A native
connector is different: its visible content is one DrawingML line and the
wire model already carries `LineOpacityThousandthPercent`. Projection already
reports that channel as `stroke.opacity`.

## Goals / Non-Goals

**Goals:**

- Make connector `compositing.opacity` validate and compile with the same
  multiplier semantics used by compound shape opacity.
- Preserve an existing stroke/color alpha by multiplying it, rather than
  replacing it.
- Keep the public PPJ output canonical and deterministic: the effective
  connector alpha is reported as `stroke.opacity`.
- Prove the path with one small NativeAOT compile-to-projection test.

**Non-Goals:**

- Implementing blend modes, group/layer isolation, clip or mask closure.
- Inventing a second native opacity channel or changing protobuf fields.
- Reconstructing an authored `compositing` declaration from an arbitrary
  imported connector. Source-bound connector edits continue to use the
  existing canonical `stroke.opacity` owner; a separate compositor-vs-paint
  split cannot be recovered from one native line alpha.

## Decisions

### 1. Lower to the existing connector line alpha

When `compositing.opacity` is below one, multiply it with the connector's
existing line alpha, defaulting that alpha to one when the stroke has no
explicit opacity. This is the same effective-alpha rule as shape branches and
uses the already-presence-aware connector field.

Alternative: add a second connector opacity field to the wire or emit a
separate overlay shape. Rejected because neither represents a distinct
PowerPoint owner and both would complicate the contract without preserving
more information.

### 2. Keep projection canonical

Projection emits the effective value as `stroke.opacity`, as it already does
for native connector line alpha. The embedded authored PPJ snapshot still
recovers the original `compositing` declaration when exact authored recovery
is requested; ordinary projection does not claim to recover a paint alpha and
an element multiplier separately.

### 3. Retain explicit fail-closed compositor boundaries

Connector opacity shares the existing normal-only checks. Non-normal blend,
true isolation, and non-empty clip stacks remain validator/compiler failures.
No lossy fallback is introduced for those semantics.

## Risks / Trade-offs

- [A connector's authored stroke opacity and compositor opacity collapse to one
  native alpha] -> Document and test the effective multiplier and expose the
  canonical `stroke.opacity` result; preserve the exact source only through
  the existing embedded PPJ snapshot.
- [Future connector paint/effect owners could be added without updating this
  rule] -> Keep the native owner statement in the PPJ reference and revisit
  the matrix in a separate change when a second owner is proven.

## Migration Plan

Existing connector programs are unchanged. Programs that add normal
`compositing.opacity` now compile; existing unsupported compositor fields keep
their current fail-closed diagnostics. Rollback is limited to removing the
connector case and its focused test/docs.
