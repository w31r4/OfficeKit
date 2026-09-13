## Context

The imported presentation theme profile already proves one shared ThemePart and six strict direct RGB accent nodes. The projection exposes all six values and the source-bound compiler already owns accent1 through accent5 through separate field-qualified capabilities. Accent6 can reuse the same repeated six-slot artifact and in-place ThemePart writer boundary.

## Goals / Non-Goals

**Goals:**

- Add an explicit `setThemeAccent6Color` owner for `accentColors.accent6`.
- Reuse the existing strict six-slot source ownership and one-theme-field-per-compile rule.
- Change only the existing accent6 RGB value and preserve no-op bytes, package residuals, and second projection.

**Non-Goals:**

- Editing multiple accent slots in one operation or widening the theme graph beyond the six direct accent nodes.
- Creating missing nodes or accepting scheme colors, alpha, transforms, multiple ThemeParts, or inherited theme state.
- Changing the wire contract, protocol version, color-role model, effect scheme, font scheme, or host rendering behavior.

## Decisions

1. **Reuse the canonical six-slot projection.** The importer and projector already retain all six direct RGB observations. The new capability targets index five; accent1 through accent5 keep their existing capabilities.
2. **Add a separate capability name.** `setThemeAccent6Color` makes authority explicit and lets semantic validation distinguish the field from the existing accent owners. A generic capability would make field ownership ambiguous.
3. **Keep one-theme-field-per-compile enforcement.** Accent6 participates in the existing single-slot mutation rule rather than introducing a second mutation path.
4. **Patch the existing owner in place.** The native writer updates only `a:accent6Color/a:srgbClr/@val` after re-reading the six direct RGB nodes. It never creates a missing node or serializes the whole theme graph.
5. **Use the existing repeated wire field.** No protobuf or protocol change is needed. Authored and imported paths remain compatible while the source-bound capability remains field-qualified.

## Risks / Trade-offs

- [Risk] Themes using scheme colors or transforms are common. → Keep them source-owned and do not issue the capability.
- [Risk] The projection shows observed values for all six slots while only the direct RGB topology is writable. → Compiler equality checks reject changes outside the requested slot and the native reference lists field-qualified authority.
- [Risk] A new RGB value can render differently across hosts. → The regression proves XML/source preservation and second projection only; it does not claim host PowerPoint appearance.

## Migration Plan

No migration is required. Existing source-bound programs remain valid. Reverting this change removes only the accent6 capability and writer branch.
