## Purpose

Expose one existing direct accent1 saturation modulation as a complete PPJ field while preserving the rest of the imported presentation theme as opaque source-owned data.

## ADDED Requirements

### Requirement: Source-bound accent1 saturation modulation is projected and edited as one bounded field

A source-bound PPJ projection MUST expose `design.theme.accentTransforms.accent1.satMod` as a fraction from `0` through `1` when the canonical imported presentation has one shared ThemePart and an accent1 color scheme entry whose direct RGB leaf has exactly one direct `a:satMod` child with an integer value from `0` through `100000`. The projection MUST issue `setThemeAccent1SatMod` for `accentTransforms.accent1.satMod`. A source-preserving compile MAY change only that existing `a:satMod/@val` token, mapping the fraction to the nearest thousandth-percent integer in the same range, and a fresh projection MUST return the edited fraction.

#### Scenario: Strict direct satMod is projected

- **WHEN** an imported presentation has one canonical ThemePart and `accent1Color/srgbClr` contains exactly one direct `a:satMod val="35000"` with no extra attributes or descendants
- **THEN** the source-bound PPJ contains `accentTransforms.accent1.satMod` equal to `0.35` and a capability whose operation is `setThemeAccent1SatMod` and whose field is `accentTransforms.accent1.satMod`

#### Scenario: Editing satMod changes only its owning token

- **WHEN** a projected program changes `accentTransforms.accent1.satMod` to `0.4` and preserves the issued capability
- **THEN** compilation succeeds, changes only the owning ThemePart's `a:satMod/@val` to `40000`, leaves all other package parts and accent1 RGB data byte-identical, and re-projection returns `0.4`

#### Scenario: No-op satMod preserves the package

- **WHEN** a projected program is compiled without changing `accentTransforms.accent1.satMod`
- **THEN** compilation succeeds with no changed parts and the original package bytes

#### Scenario: Unsupported topology fails closed

- **WHEN** the ThemePart is missing or ambiguous, the accent1 RGB leaf is absent or transformed, `a:satMod` is missing, has extra attributes or descendants, has an out-of-range value, or has unsupported transform siblings
- **THEN** projection omits the satMod field and capability, and a compile request cannot create or replace that source-owned topology

#### Scenario: Deletion, combined edits, and capability tampering fail closed

- **WHEN** a source-bound program deletes satMod, changes satMod together with another theme-owned field, or changes the capability operation or field list
- **THEN** compilation fails with an unsupported source-bound edit diagnostic and does not produce a modified package
