## Context

The authored compiler already lowers `design.theme.name` into `PresentationThemeArtifact.Name`, while imported projections currently use a placeholder name and have no theme-level native reference. Source-preserving export already validates the complete source package and reports changed package parts.

## Goals / Non-Goals

**Goals:**

- Expose one stable PPJ field for an existing presentation theme name.
- Restrict the owner to one unique ThemePart shared by every slide master.
- Change only the unqualified `a:theme/@name` attribute and retain all theme children and attributes.
- Keep the source package hash as the source binding and use a stable theme capability object hash.

**Non-Goals:**

- Editing font schemes, color schemes, effect schemes, inheritance, or theme relationships.
- Creating a missing theme name, merging multiple themes, or rewriting a ThemePart.
- Claiming PowerPoint host rendering or complete F-15 support.

## Decisions

1. **Use `design.theme.name` as the first source-bound theme field.** The field already exists in the PPJ schema and authored wire model, so this increment adds the smallest source-preserving owner.
2. **Require one shared existing ThemePart and an existing non-empty name.** If masters have no theme, different theme parts, or a missing name, the projection remains source-owned and no capability is issued. This avoids guessing which theme controls the deck.
3. **Bind the capability at the theme object.** A stable `theme` scope and object hash are used for the capability; the package source SHA-256 still binds the exact source bytes. The requested native reference must otherwise match the fresh projection.
4. **Patch the SDK theme root in place.** The writer sets only the existing root `Name` value, saves that ThemePart, records its part path, and leaves all descendants untouched. A post-write check confirms the requested name.

## Risks / Trade-offs

- [Risk] Theme names are metadata and may not change visible rendering. → The field is documented as a narrow F-15 metadata owner and host visual acceptance remains out of scope.
- [Risk] A producer can use multiple master-specific themes. → Such input receives no theme capability and remains source-preserved.
- [Risk] A caller removes the optional PPJ name. → The compiler treats that as an unsupported deletion because this increment only owns an existing name token.

## Migration Plan

No migration is required. Existing imported PPJ without a theme native reference remains valid; fresh projections of canonical single-theme files gain the optional reference. Reverting the commit restores the previous placeholder projection and fail-closed behavior.
