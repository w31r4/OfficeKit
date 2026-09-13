## Context

The committed dark1, light1, dark2, light2, and hyperlink increments import six strict direct RGB color-role leaves into the existing `PresentationThemeArtifact`, project the complete six-role observation, and authorize independent field-qualified source-bound edits. The final followed-hyperlink role must reuse that same all-six observation boundary.

## Goals / Non-Goals

**Goals:**

- Add `setThemeColorRoleFollowedHyperlink` for `design.theme.colorRoles.followedHyperlink` while keeping the existing color-role capabilities intact.
- Patch only an existing direct `a:folHlink/a:srgbClr/@val` token in the canonical shared ThemePart.
- Prove source-byte, changed-part, XML, and authority preservation with a small native fixture.

**Non-Goals:**

- Editing arbitrary theme roles or theme graph state in this change.
- Creating missing nodes or accepting scheme colors, alpha, transforms, inherited state, multiple ThemeParts, or arbitrary theme XML.
- Changing protobuf fields, protocol version, authored color-role lowering, effect schemes, font schemes, or host rendering behavior.

## Decisions

1. **Reuse the existing six-role artifact.** `PresentationThemeArtifact.FollowedHyperlinkRgb` is already present for authored lowering and is populated by the strict imported-role reader. This keeps the wire contract stable.
2. **Issue an independent field capability.** Add `setThemeColorRoleFollowedHyperlink` with only `colorRoles.followedHyperlink`; retain existing capabilities so a projection remains truthful about all independently owned fields.
3. **Keep one theme field per compile.** The compiler's existing mutation guard rejects changing followed-hyperlink together with other roles, accents, fonts, or any other theme field.
4. **Patch the existing XML token.** The writer rereads the canonical ThemePart, verifies the direct RGB followed-hyperlink topology, and changes only `a:folHlink/a:srgbClr/@val`. It never synthesizes a missing node or serializes unsupported descendants.
5. **Preserve the strict profile.** Alpha, transforms, system colors, extra attributes, and ambiguous ownership remain source-owned or fail closed. A broader role graph would require a separate spec and test boundary.

## Risks / Trade-offs

- Many themes use a system or scheme color for followed hyperlinks, so those imports intentionally receive no color-role capabilities; this is safer than claiming an edit that cannot preserve the source topology.
- The complete six-role observation is required by the current PPJ schema even though the bounded roles have independent capabilities; unsupported state remains observable but source-owned.
- The focused regression proves package/XML preservation and reprojection, not PowerPoint host color management or visual fidelity.

## Migration Plan

No migration is required. Existing PPJ programs remain valid. Rolling back removes only the followed-hyperlink observation authority and writer branch; authored theme color roles and existing source-bound owners remain available.
