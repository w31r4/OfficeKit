## Purpose

Expose one direct imported accent1 green-channel modulation as a complete PPJ field while preserving every other part of the presentation theme as source-owned data.

## ADDED Requirements

### Requirement: Source-bound accent1 green modulation is projected and edited as one bounded field

A source-bound PPJ projection MUST expose `design.theme.accentTransforms.accent1.greenMod` as a fraction from `0` through `1` when the canonical imported presentation has one shared ThemePart and an accent1 color scheme entry whose direct RGB leaf has exactly one direct `a:greenMod` child with an integer value from `0` through `100000`. The projection MUST issue `setThemeAccent1GreenMod` for `accentTransforms.accent1.greenMod`. A source-preserving compile MAY change only that existing `a:greenMod/@val` token, mapping the fraction to the nearest thousandth-percent integer in the same range, and a fresh projection MUST return the edited fraction.

#### Scenario: Strict direct greenMod is projected

- **WHEN** an imported presentation has one canonical ThemePart and `accent1Color/srgbClr` contains exactly one direct `a:greenMod val="35000"` with no extra attributes or descendants
- **THEN** the source-bound PPJ contains `accentTransforms.accent1.greenMod` equal to `0.35` and a capability whose operation is `setThemeAccent1GreenMod` and whose field is `accentTransforms.accent1.greenMod`

#### Scenario: Editing greenMod changes only its owning token

- **WHEN** a projected program changes `accentTransforms.accent1.greenMod` to `0.4` and preserves the issued capability
- **THEN** compilation succeeds, changes only the owning ThemePart's `a:greenMod/@val` to `40000`, leaves all other package parts and accent1 RGB data byte-identical, and re-projection returns `0.4`

#### Scenario: No-op greenMod preserves the package

- **WHEN** a projected program is compiled without changing `accentTransforms.accent1.greenMod`
- **THEN** compilation succeeds with no changed parts and the original package bytes

#### Scenario: Unsupported topology fails closed

- **WHEN** the ThemePart is missing or ambiguous, the accent1 RGB leaf is absent or transformed, `a:greenMod` is missing, has extra attributes or descendants, has an out-of-range value, or has unsupported transform siblings
- **THEN** projection omits the greenMod field and capability, and a compile request cannot create or replace that source-owned topology

#### Scenario: Deletion, combined edits, and capability tampering fail closed

- **WHEN** a source-bound program deletes greenMod, changes greenMod together with another theme-owned field, or changes the capability operation or field list
- **THEN** compilation fails with an unsupported source-bound edit diagnostic and does not produce a modified package
