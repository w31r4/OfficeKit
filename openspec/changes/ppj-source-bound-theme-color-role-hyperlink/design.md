## Context

The committed dark1, light1, dark2, and light2 increments import six strict direct RGB color-role leaves into the existing `PresentationThemeArtifact`, project the complete six-role observation, and authorize independent field-qualified source-bound edits. The generated PPJ schema requires all six `colorRoles` properties, so a hyperlink capability must reuse that same all-six observation boundary.

## Goals / Non-Goals

**Goals:**

- Add `setThemeColorRoleHyperlink` for `design.theme.colorRoles.hyperlink` while keeping the existing color-role capabilities intact.
- Patch only an existing direct `a:hlink/a:srgbClr/@val` token in the canonical shared ThemePart.
- Prove source-byte, changed-part, XML, and authority preservation with a small native fixture.

**Non-Goals:**

- Editing followedHyperlink in this change.
- Creating missing nodes or accepting scheme colors, alpha, transforms, inherited state, multiple ThemeParts, or arbitrary theme XML.
- Changing protobuf fields, protocol version, authored color-role lowering, effect schemes, font schemes, or host rendering behavior.

## Decisions

1. **Reuse the existing six-role artifact.** `PresentationThemeArtifact.hyperlink_rgb` is already present for authored lowering and is populated by the strict imported-role reader. This keeps the wire contract stable.
2. **Issue an independent field capability.** Add `setThemeColorRoleHyperlink` with only `colorRoles.hyperlink`; retain existing capabilities so a projection remains truthful about all independently owned fields.
3. **Keep one theme field per compile.** The compiler's existing mutation guard rejects changing hyperlink together with other roles, accents, fonts, or any other theme field.
4. **Patch the existing XML token.** The writer rereads the canonical ThemePart, verifies the direct RGB hyperlink topology, and changes only `a:hlink/a:srgbClr/@val`. It never synthesizes a missing node or serializes unsupported descendants.
5. **Preserve the strict profile.** Alpha, transforms, system colors, extra attributes, and ambiguous ownership remain source-owned or fail closed. A broader role graph would require a separate spec and test boundary.

## Risks / Trade-offs

- Many themes use a system or scheme color for hyperlink, so those imports intentionally receive no color-role capabilities; this is safer than claiming an RGB edit that cannot preserve the source topology.
- The complete six-role observation is required by the current PPJ schema even though only the bounded roles have independent capabilities; unsupported roles remain observable but source-owned.
- The focused regression proves package/XML preservation and reprojection, not PowerPoint host color management or visual fidelity.

## Migration Plan

No migration is required. Existing PPJ programs remain valid. Rolling back removes only the hyperlink observation authority and writer branch; authored theme color roles and existing source-bound owners remain available.
