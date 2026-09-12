## Context

The committed PPJ text projection parses a strict direct rich-text outer shadow and already carries its RGB color in the semantic shadow object. Geometry leaves use the same run-index/source-bound path. This change exposes the existing direct `srgbClr/@val` as one independent native leaf without rebuilding the effect list or touching unrelated package members.

## Goals / Non-Goals

**Goals:**

- Carry an existing canonical direct `outerShdw/srgbClr/@val` through native projection, PPJ schema/registry, edit proof, token patching, and second projection.
- Preserve the existing blur, distance, direction, opacity, and source-owned color topology behavior.
- Exercise the normal source-bound footprint and fail-closed boundaries with a focused codec fixture.

**Non-Goals:**

- Adding direct run theme-color or opacity leaves in this increment.
- Creating or deleting a missing color child, converting theme colors, accepting transformed or compound effect graphs, or proving PowerPoint rendering.
- Changing the wire protocol version or implementing a general effect editor.

## Decisions

1. **Use a sibling native leaf.** `textShadowColorRgb` maps directly to `run.style.shadow.color`, so callers can edit RGB independently while existing geometry fields remain unchanged.

2. **Require one direct RGB child.** The existing strict `PptxShadowCodec` profile already rejects multiple effect-list children, unsupported attributes, unknown descendants, missing colors, malformed RGB tokens, and theme-color-only owners. The new leaf is emitted only when that profile exposes `ColorRgb`.

3. **Keep canonical six-hex tokens.** Projection stores the native token as uppercase and the typed PPJ value as a `#`-prefixed lowercase color. Source-bound proof normalizes both expected and requested values, then patches only the six-character `val` token.

4. **Reuse the run-index/source-bound patch path.** The text leaf index identifies the owning run. Proof re-reads the strict shadow and verifies the current direct `srgbClr/@val`; patching preserves alpha, geometry, run topology, and non-target OPC parts by construction.

## Risks / Trade-offs

- A source with a theme color remains source-owned for this RGB leaf; callers must use a separate future field if theme edits are needed.
- A noncanonical or stale token cannot pass proof because both the strict parser and raw-token precondition require six hexadecimal characters and the planned value must change.
- A future effect-list feature that changes the run index is handled by the existing source hash and ownership proof; the edit fails closed instead of guessing.
