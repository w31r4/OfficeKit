## Purpose

Expose one existing direct accent1 red-channel offset as a complete PPJ field while preserving the rest of the imported presentation theme as opaque source-owned data.

## ADDED Requirements

### Requirement: Source-bound accent1 red offset is projected and edited as one bounded field

A source-bound PPJ projection MUST expose `design.theme.accentTransforms.accent1.redOff` as a signed fraction from `-1` through `1` when the canonical imported presentation has one shared ThemePart and an accent1 color scheme entry whose direct RGB leaf has exactly one direct `a:redOff` child with an integer value from `-100000` through `100000`. The projection MUST issue `setThemeAccent1RedOff` for `accentTransforms.accent1.redOff`. A source-preserving compile MAY change only that existing `a:redOff/@val` token, mapping the signed fraction to the nearest thousandth-percent integer in the same range, and a fresh projection MUST return the edited fraction.

#### Scenario: Strict direct redOff is projected

- **WHEN** an imported presentation has one canonical ThemePart and `accent1Color/srgbClr` contains exactly one direct `a:redOff val="-25000"` with no extra attributes or descendants
- **THEN** the source-bound PPJ contains `accentTransforms.accent1.redOff` equal to `-0.25` and a capability whose operation is `setThemeAccent1RedOff` and whose field is `accentTransforms.accent1.redOff`

#### Scenario: Editing redOff changes only its owning token

- **WHEN** a projected program changes `accentTransforms.accent1.redOff` to `0.4` and preserves the issued capability
- **THEN** compilation succeeds, changes only the owning ThemePart's `a:redOff/@val` to `40000`, leaves all other package parts and accent1 RGB data byte-identical, and re-projection returns `0.4`

#### Scenario: No-op redOff preserves the package

- **WHEN** a projected program is compiled without changing `accentTransforms.accent1.redOff`
- **THEN** compilation succeeds with no changed parts and the original package bytes

#### Scenario: Unsupported topology fails closed

- **WHEN** the ThemePart is missing or ambiguous, the accent1 RGB leaf is absent or transformed, `a:redOff` is missing, has extra attributes or descendants, has an out-of-range value, or has unsupported transform siblings
- **THEN** projection omits the redOff field and capability, and a compile request cannot create or replace that source-owned topology

#### Scenario: Deletion, combined edits, and capability tampering fail closed

- **WHEN** a source-bound program deletes redOff, changes redOff together with another theme-owned field, or changes the capability operation or field list
- **THEN** compilation fails with an unsupported source-bound edit diagnostic and does not produce a modified package
