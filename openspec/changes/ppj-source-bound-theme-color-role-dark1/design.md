## Context

The imported theme reader already resolves one canonical ThemePart and the native artifact already has an optional `dark1_rgb` field used by authored lowering. The source-bound projection currently leaves all `colorRoles` opaque. A direct six-digit RGB dark1 node can reuse the existing artifact and writer boundary.

## Goals / Non-Goals

**Goals:**

- Add an explicit `setThemeColorRoleDark1` owner for `design.theme.colorRoles.dark1`.
- Accept only one shared ThemePart with an existing direct `a:dk1/a:srgbClr` leaf.
- Patch only that leaf and preserve the rest of the package and theme graph.

**Non-Goals:**

- Editing light1, dark2, light2, hyperlink, or followedHyperlink in this change.
- Creating missing color nodes or accepting scheme colors, alpha, transforms, multiple ThemeParts, or inherited state.
- Changing the wire contract, protocol version, authored color-role semantics, effect scheme, font scheme, or host rendering behavior.

## Decisions

1. Reuse `PresentationThemeArtifact.dark1_rgb` for both imported observation and source-bound mutation.
2. Add a field-qualified capability so authority is explicit and semantic validation can reject edits to the other roles.
3. Keep the existing one-theme-field-per-compile rule; dark1 is one independent color-role field.
4. Re-read the canonical ThemePart and replace only `a:dk1/a:srgbClr/@val`; never synthesize or serialize an unmodeled theme graph.
5. Keep the strict bare RGB profile. Alpha and any child transforms remain source-owned.

## Risks / Trade-offs

- Themes frequently use system or scheme colors for dark1. Such themes do not receive the capability.
- A projected role is an observed source value, not a claim about the full theme cascade or host appearance.
- The regression proves package/XML/source preservation and reprojection, not Windows PowerPoint rendering.

## Migration Plan

No migration is required. Existing PPJ programs remain valid. Reverting this change removes only the dark1 observation, capability, and writer branch.
