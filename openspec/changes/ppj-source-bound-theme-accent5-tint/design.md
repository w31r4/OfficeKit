## Context

The existing source-preserving theme path recognizes strict direct transform
leaves for accent1 through accent4 and patches the existing DrawingML token.
This change applies the same ownership boundary to accent5 tint without
broadening the theme graph or changing authored six-role lowering.

## Goals / Non-Goals

**Goals:**

- Recognize only a canonical shared ThemePart whose accent5 color is a direct
  six-digit RGB leaf with one direct `a:tint/@val` child.
- Expose a fraction and a field-qualified capability, then patch that one
  existing token with a bounded conversion.
- Preserve all unrelated theme roles, transforms, package parts, and no-op
  bytes.

**Non-Goals:**

- Accent5 transforms other than `tint`, all accent6 transforms, inherited or
  effect theme state, or arbitrary theme XML editing.
- Creating or deleting a missing tint node, changing multiple theme fields, or
  claiming PowerPoint host color rendering.

## Decisions

1. **Use the existing native transform field.**
   `PresentationThemeColorTransform.TintThousandth` already represents the
   DrawingML 1/100000 tint integer and the authored compiler already emits it.
   The importer records a role of `accent5`; the projector converts the native
   value by dividing by 100000, and the source-bound compiler rounds by
   `Math.Round(value * 100000d, MidpointRounding.AwayFromZero)`.

2. **Keep the source owner strict and role-specific.**
   The importer helper requires exactly one shared ThemePart, a direct RGB
   accent5 child, no extra attributes or descendants, and a single direct
   `a:tint`. The source writer patches `Accent5Color` only when the authored
   transform is the one existing accent5 tint and rejects all other transform
   flags.

3. **Keep capability authority field-qualified.**
   The operation `setThemeAccent5Tint` is accepted only with
   `accentTransforms.accent5.tint` in the source-issued capability. This keeps
   forged operations or retargeted fields fail-closed while retaining the
   existing one-theme-field-per-compile rule.

## Risks / Trade-offs

- **Risk:** A theme may contain transforms or inheritance that look similar but
  do not have a stable local owner. → **Mitigation:** omit capability and
  retain original source topology whenever strict shape is not met.
- **Risk:** Fraction-to-integer conversion could drift at half 1/100000.
  → **Mitigation:** use explicit away-from-zero rounding and reproject native
  integer in focused regression.
- **Risk:** Writer could touch unrelated package data. → **Mitigation:** assert
  canonical ThemePart is only changed part and compare every other ZIP part
  byte-for-byte.
