## Purpose

Expose one complete, field-qualified source-bound PPJ owner for an existing presentation theme followed-hyperlink color while preserving the imported theme graph and the existing color-role owners.

## ADDED Requirements

### Requirement: Canonical followed-hyperlink color-role projection

A source-derived PPJ projection SHALL expose `design.theme.colorRoles` from six existing direct `a:clrScheme` RGB leaves (`a:dk1`, `a:lt1`, `a:dk2`, `a:lt2`, `a:hlink`, and `a:folHlink`) when every slide master resolves to one unique ThemePart and all six nodes are six-digit RGB leaves with no descendants or extra attributes. The projection SHALL issue `setThemeColorRoleFollowedHyperlink` with only `colorRoles.followedHyperlink` in the theme native reference, while retaining any already supported capabilities for other owned roles.

#### Scenario: Existing direct followed-hyperlink color is projected with authority

- **WHEN** an imported presentation has one shared ThemePart whose six bounded color roles are direct six-digit RGB leaves
- **THEN** `design.theme.colorRoles.followedHyperlink` equals the source followed-hyperlink value and the theme native reference issues `setThemeColorRoleFollowedHyperlink` for `colorRoles.followedHyperlink`

#### Scenario: Unsupported followed-hyperlink ownership stays source-owned

- **WHEN** the source has no unique ThemePart, a missing bounded color-role node, a scheme/system color, an alpha child, a transform, or extra topology
- **THEN** the projection omits `colorRoles` and does not issue the followed-hyperlink capability

### Requirement: Source-bound followed-hyperlink edit

A source-bound PPJ compile SHALL accept a changed existing `design.theme.colorRoles.followedHyperlink` only with the freshly projected theme native reference and its `setThemeColorRoleFollowedHyperlink` capability. The export SHALL replace only the owning ThemePart's `a:folHlink/a:srgbClr/@val`, report that ThemePart as changed, and preserve all other theme roles, accents, descendants, relationships, and package members.

#### Scenario: Followed-hyperlink edit changes one theme part

- **WHEN** a caller changes the projected followed-hyperlink value and retains the issued native reference
- **THEN** compilation succeeds, reports only the canonical ThemePart as changed, and a second projection returns the new followed-hyperlink value

#### Scenario: Other roles or authority are rejected

- **WHEN** a caller deletes followed-hyperlink, adds or changes another color role in the same object, changes another theme field in the same compile, or tampers with the capability
- **THEN** compilation fails without rewriting the source package

#### Scenario: No-op preserves source bytes

- **WHEN** a caller compiles the unchanged projected PPJ against the exact source package
- **THEN** output bytes and changed-part list are identical to the source and empty respectively
