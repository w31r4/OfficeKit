## Context

The authored compiler already stores `PresentationThemeArtifact.accent_rgb` and lowers it to `a:accent1Color` through `a:accent6Color`. Imported projections currently expose only a placeholder theme name and reject `accentColors` edits because the imported ThemePart is source-owned. The existing source-preserving writer already has changed-part and opaque-content proofs.

## Goals / Non-Goals

**Goals:**

- Read six direct RGB accent colors from one canonical shared ThemePart.
- Reuse the existing repeated wire field and bind one theme-level capability to `accentColors.accent1`.
- Patch the existing accent1 `val` attribute in place while preserving the remaining theme graph.

**Non-Goals:**

- Editing accent2 through accent6, dark/light or hyperlink roles, transforms, effect schemes, or font schemes in this increment.
- Creating missing color nodes, accepting `schemeClr`, alpha, transforms, or choosing among multiple ThemeParts.
- Modeling theme inheritance, color management, or host rendering.

## Decisions

1. **Require one shared ThemePart plus six strict direct RGB accents.** The PPJ object requires all six accent fields, so the importer exposes the object only when every accent is an existing `srgbClr` with a six-digit `val` and no descendants. A missing, transformed, alpha-bearing, or ambiguous source graph receives no accent capability.
2. **Keep one theme native reference with a field-qualified capability.** The existing theme-level native reference carries `setThemeAccent1Color`; `accentColors.accent1` is the only newly writable field. The other observed accents remain equality-checked source state.
3. **Use the existing repeated artifact field.** Import fills six `AccentRgb` entries, compilation changes index zero, and the PPTX writer applies the values after re-proving the canonical ThemePart. No protobuf field or protocol version changes.
4. **Write only the existing accent1 attribute.** The writer updates `a:accent1Color/a:srgbClr/@val`, saves the ThemePart, records its path and replacement hash, and lets the existing opaque graph guard prove all untouched parts byte-for-byte.
5. **Normalize comparison without rewriting no-ops.** Imported values are projected in uppercase, while the writer compares case-insensitively and only assigns a changed value. Lowercase source bytes therefore remain byte-identical on a no-op.

## Risks / Trade-offs

- [Risk] Common Office themes use `sysClr`, `schemeClr`, alpha, or color transforms. → Keep those themes source-owned until a matching bounded profile exists.
- [Risk] The PPJ object exposes five unowned accents beside one writable slot. → Compiler equality checks reject every change except `accentColors.accent1`, and the capability lists only that path.
- [Risk] A valid RGB value changes host appearance differently across applications. → The field records native XML and does not claim PowerPoint rendering parity.

## Migration Plan

No migration is required. Existing source-bound programs without the optional capability remain valid. Reverting the change restores the previous opaque `accentColors` behavior.
