## Context

The source-bound accent palette currently accepts only six-digit direct RGB leaves for accent2 through accent6. Authored PPJ already accepts six- or eight-digit theme colors and lowers the trailing byte to `a:alpha`; the imported profile does not yet preserve that child for accent2.

## Goals / Non-Goals

**Goals:**

- Read `accent2` as `#RRGGBBAA` when its direct `srgbClr` has exactly one canonical `a:alpha` child.
- Keep the existing six accent object and `setThemeAccent2Color` field authority.
- Update the existing RGB and alpha attributes together, with a single ThemePart changed-part receipt and second-projection proof.

**Non-Goals:**

- Adding alpha ownership to accent3 through accent6 or the six non-accent color roles in this increment.
- Creating missing alpha nodes, changing alpha presence, accepting transforms/scheme colors, or modeling theme inheritance/effect schemes.

## Decisions

1. **Use the existing string field.** `PresentationThemeArtifact.accent_rgb[1]` already carries authored RGB/RGBA and avoids a new protobuf field or wire version.
2. **Require a canonical 8-bit alpha.** Only `a:alpha/@val` values produced by the existing `#RRGGBBAA` lowering are projected, so a second projection can reproduce the exact byte value and no-op compilation remains byte-identical.
3. **Keep owner topology strict.** The RGB leaf must have only `val` plus one alpha child, and the alpha child must be pre-existing. Requests must preserve whether alpha is present; no new child or deletion is inferred.
4. **Patch one field owner.** The writer compares RGB and alpha separately, updates only those two attributes under accent2, and leaves the other five accents and all theme descendants source-owned.

## Risks / Trade-offs

- Office files may use arbitrary thousandths alpha or additional transforms; those remain opaque until a matching canonical PPJ representation exists.
- The existing capability name still covers the RGB/alpha pair; callers must retain the projected field and hash-bound native reference.
