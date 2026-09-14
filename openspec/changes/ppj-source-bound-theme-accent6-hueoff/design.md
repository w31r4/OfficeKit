## Context

The existing source-preserving theme path recognizes one direct accent1 hue
transform and direct accent6 tint, shade, luminance, alpha, saturation, red,
blue, and hue modulation leaves at a time, then patches the existing DrawingML
token. This change applies the same ownership boundary to accent6 hue offset
without broadening the theme graph or changing authored six-role lowering.

## Goals / Non-Goals

**Goals:**

- Recognize only a canonical shared ThemePart whose accent6 color is a direct
  six-digit RGB leaf with one direct `a:hueOff/@val` child.
- Expose a degree offset and a field-qualified capability, then patch that one
  existing token with a bounded conversion.
- Preserve all unrelated theme roles, transforms, package parts, and no-op
  bytes.

**Non-Goals:**

- Accent6 transforms other than `hueOff`, transforms on accent6 through
  accent6, inherited or effect theme state, or arbitrary theme XML editing.
- Creating or deleting a missing hueOff node, changing multiple theme fields,
  or claiming PowerPoint host color rendering.

## Decisions

1. **Use the existing native transform field.**
   `PresentationThemeColorTransform.HueOffsetAngleThousandth` already represents
   the DrawingML 1/60000-degree signed offset integer and the authored compiler
   already emits it. The importer records a role of `accent6`; the projector
   converts the native value by dividing by 60000, and the source-bound compiler
   rounds by `Math.Round(value * 60000d, MidpointRounding.AwayFromZero)`.

2. **Keep the source owner strict and role-specific.**
   The importer helper requires exactly one shared ThemePart, a direct RGB
   accent6 child, no extra attributes or descendants, and a single direct
   `a:hueOff`. The source writer patches `Accent6Color` only when the authored
   transform is the one existing accent6 hueOff and rejects all other transform
   flags.

3. **Keep capability authority field-qualified.**
   The operation `setThemeAccent6HueOff` is accepted only with
   `accentTransforms.accent6.hueOff` in the source-issued capability. This
   keeps forged operations or retargeted fields fail-closed while retaining the
   existing one-theme-field-per-compile rule.

## Risks / Trade-offs

- **Risk:** A theme may contain transforms or inheritance that look similar but
  do not have a stable local owner. → **Mitigation:** omit the capability and
  retain the original source topology whenever the strict shape is not met.
- **Risk:** Degree-to-integer conversion could drift at half 1/60000 degree.
  → **Mitigation:** use explicit away-from-zero rounding and reproject the
  native integer in the focused regression.
- **Risk:** A writer change could touch unrelated package data. → **Mitigation:**
  assert the canonical ThemePart is the only changed part and compare every
  other ZIP part byte-for-byte.
