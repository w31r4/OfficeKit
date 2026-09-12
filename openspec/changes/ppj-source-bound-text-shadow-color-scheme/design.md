## Context

The shared shadow model and authored compiler already write direct RGB or theme
colors. The strict imported direct-run shadow parser also retains
`PresentationShadow.ColorScheme`; only the native leaf and source-bound patch
surface currently omit the theme form. This increment mirrors the existing
RGB leaf and default-text theme-color path.

## Goals / Non-Goals

**Goals:**

- Carry one existing bare direct `outerShdw/schemeClr/@val` through projection,
  PPJ metadata, proof/readback, token patching, and second projection.
- Preserve the existing shadow geometry, alpha, topology, and source ownership
  rules.
- Exercise a focused theme-token fixture and fail-closed boundaries.

**Non-Goals:**

- Creating or deleting a missing color child or converting RGB to theme color.
- Accepting tint/shade or other color transforms, multiple color children,
  transformed/compound effect graphs, or proving PowerPoint rendering.
- Changing the wire protocol version or implementing a general effect editor.

## Decisions

1. **Use a sibling native leaf.** `textShadowColorScheme` maps to the existing
   `run.style.shadow.color` theme-token form while the RGB leaf remains separate.

2. **Require one bare scheme child.** The strict shadow parser already rejects
   duplicate colors, malformed tokens, color transforms, unsupported attributes,
   unknown descendants, missing geometry, and compound effect graphs. The new
   leaf is emitted only when that parser exposes `ColorScheme`.

3. **Reuse the run-index source-bound patch path.** The text leaf index identifies
   the owning run. Proof normalizes expected and requested theme tokens and the
   patch replaces only `schemeClr/@val`, preserving all other XML and OPC parts.

## Risks / Trade-offs

- A transformed theme color remains source-owned; callers cannot request a
  tint/shade edit through this bounded leaf.
- Noncanonical or stale tokens fail proof instead of being rewritten broadly.
- Focused codec evidence does not establish host PowerPoint display behavior.

## Migration Plan

No data migration is required. Existing PPJ documents remain valid; imported
runs gain this leaf only when they satisfy the bounded profile. The change can
be rolled back as one commit without changing the wire version.
