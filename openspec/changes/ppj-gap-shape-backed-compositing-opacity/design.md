## Context

See `proposal.md` for the gap. `icon` is lowered to a custom editable shape
with catalog geometry, while `placeholder` is lowered to a textbox shape with
placeholder identity. The shared authored compositor already handles shape
fill, line, shadow, and explicit text paint alpha.

## Goals / Non-Goals

**Goals:**

- Make the existing `compositing.opacity` contract match the two existing
  shape-backed authored element types.
- Keep alpha multiplication and unsupported compositor diagnostics identical
  to ordinary shape/text lowering.
- Prove both paths with one compact compile/project test.

**Non-Goals:**

- Adding a new native or wire opacity channel.
- Reconstructing `iconName` or placeholder compositor declarations from
  ordinary imported shapes.
- Adding opacity to groups, charts, tables, media, or unsupported effect
  graphs; those require separate native-owner decisions.

## Decisions

### 1. Expand the semantic type allow-list only

The compiler's `BuildElement` already emits `PresentationShape` for icons and
placeholders, so the implementation only needs to admit their model types in
the existing opacity validation. The existing `ApplyAuthoredCompositing`
shape branch then owns the native lowering.

Alternative: duplicate icon/placeholder-specific opacity code. Rejected
because it would create a second alpha rule and could diverge from shape text
paint handling.

### 2. Keep ordinary projection truthful

The test removes the embedded PPJ snapshot and checks the resulting native
shape paint. Ordinary projection may represent an icon as a shape, as it does
today, and may retain a placeholder's native identity; this change does not
invent an authored-only field in the imported output.

## Risks / Trade-offs

- [Placeholder text with inherited or unsupported paint may not have a safe
  alpha owner] -> Reuse the existing compound-shape checks, which fail closed
  when the text paint cannot be represented.
- [Icon identity is not recoverable from arbitrary geometry] -> Keep the
  current shape-backed projection boundary and use the embedded snapshot only
  for exact authored recovery.

## Migration Plan

Existing programs are unchanged. Programs adding normal opacity to an icon or
placeholder now compile; unsupported blend/isolation/clip declarations retain
their current diagnostics. Removing the validator allow-list cleanly rolls the
slice back.
