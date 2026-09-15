## Purpose

Extend the existing field-qualified source-bound PPJ owner for an existing presentation theme dark2 color while preserving the imported theme graph and the existing dark1/light1 owners.

## ADDED Requirements

### Requirement: Canonical dark2 color-role projection

A source-derived PPJ projection SHALL expose `design.theme.colorRoles` from six existing direct `a:clrScheme` RGB leaves (`a:dk1`, `a:lt1`, `a:dk2`, `a:lt2`, `a:hlink`, and `a:folHlink`) when every slide master resolves to one unique ThemePart, dark2 is a six- or eight-digit RGB/RGBA leaf, and the other five nodes are six-digit RGB leaves with no descendants or extra attributes. The projection SHALL issue `setThemeColorRoleDark2` with only `colorRoles.dark2` in the theme native reference, while retaining any already supported capabilities for other owned roles.

#### Scenario: Existing direct dark2 is projected with authority

- **WHEN** an imported presentation has one shared ThemePart whose dark2 role is a direct six- or eight-digit RGB/RGBA leaf and whose other five roles are direct six-digit RGB leaves
- **THEN** `design.theme.colorRoles.dark2` equals the source dark2 value and the theme native reference issues `setThemeColorRoleDark2` for `colorRoles.dark2`

#### Scenario: Unsupported dark2 ownership stays source-owned

- **WHEN** the source has no unique ThemePart, a missing bounded color-role node, a scheme/system color, a non-canonical alpha child, a transform, or extra topology
- **THEN** the projection omits `colorRoles` and does not issue the dark2 capability

### Requirement: Source-bound dark2 edit

A source-bound PPJ compile SHALL accept a changed existing `design.theme.colorRoles.dark2` only with the freshly projected theme native reference and its `setThemeColorRoleDark2` capability. The export SHALL replace only the owning ThemePart's `a:dk2/a:srgbClr/@val` and its existing `a:alpha/@val`, report that ThemePart as changed, and preserve all other theme roles, accents, descendants, relationships, and package members.

#### Scenario: Dark2 edit changes one theme part

- **WHEN** a caller changes the projected dark2 value and retains the issued native reference
- **THEN** compilation succeeds, reports only the canonical ThemePart as changed, and a second projection returns the new dark2 value

#### Scenario: Other roles or authority are rejected

- **WHEN** a caller deletes dark2, adds or changes another color role in the same object, changes another theme field in the same compile, or tampers with the capability
- **THEN** compilation fails without rewriting the source package

#### Scenario: No-op preserves source bytes

- **WHEN** a caller compiles the unchanged projected PPJ against the exact source package
- **THEN** output bytes and changed-part list are identical to the source and empty respectively
