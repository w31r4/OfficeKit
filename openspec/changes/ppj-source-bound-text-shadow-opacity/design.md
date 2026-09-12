## Context

The committed PPJ text projection parses a strict direct rich-text outer
shadow and already carries its optional alpha as
`PresentationShadow.OpacityThousandthPercent`. Geometry and RGB color leaves
use the same run-index/source-bound path. This change exposes the existing
direct `a:alpha/@val` as one independent native leaf without rebuilding the
effect list or touching unrelated package members.

## Goals / Non-Goals

**Goals:**

- Carry an existing canonical direct `outerShdw/(srgbClr|schemeClr)/alpha/@val`
  through native projection, PPJ schema/registry, edit proof, token patching,
  and second projection.
- Keep the typed `run.style.shadow.opacity` value and native leaf aligned at
  0..1 semantic opacity and 0..100000 native thousandth-percent precision.
- Exercise the normal source-bound footprint and fail-closed boundaries with a
  focused codec fixture.

**Non-Goals:**

- Adding direct run theme-color or other shadow-effect leaves in this increment.
- Creating or deleting a missing alpha child, converting colors, accepting
  transformed or compound effect graphs, or proving PowerPoint rendering.
- Changing the wire protocol version or implementing a general effect editor.

## Decisions

1. **Use a sibling native leaf.** `textShadowOpacityThousandthPercent` maps
   directly to `run.style.shadow.opacity`, so callers can edit alpha
   independently while the existing geometry and color fields remain intact.

2. **Allow either direct color owner.** Alpha is independent of color
   representation, so a strict outer shadow with exactly one direct RGB or
   theme color child may expose this field when that child contains one valid
   alpha. The RGB color leaf remains limited to RGB owners.

3. **Keep canonical integer tokens.** Projection stores the native alpha as a
   canonical integer from 0 through 100000 and the typed PPJ value as the
   corresponding ratio. Source-bound proof requires both expected and
   requested values to be changed canonical integers, then splices only the
   six-digit-or-shorter alpha `val` token.

4. **Reuse the run-index/source-bound patch path.** The text leaf index
   identifies the owning run. Proof re-reads the strict shadow and verifies the
   current direct color child's alpha; patching preserves color, geometry, run
   topology, and non-target OPC parts by construction.

## Risks / Trade-offs

- A source without an existing alpha remains source-owned; this increment does
  not invent effect children or presence.
- A noncanonical, stale, or out-of-range token cannot pass proof because the
  strict parser and raw-token precondition require 0..100000 canonical values.
- A future effect-list feature that changes the run index is handled by the
  existing source hash and ownership proof; the edit fails closed instead of
  guessing.
