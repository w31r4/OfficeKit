## 1. Contract

- [x] 1.1 Extend the PPJ theme capability vocabulary and source-bound coverage entry for `design.theme.fontScheme.majorEastAsia`, then validate the OpenSpec change
- [x] 1.2 Keep the existing `PresentationThemeArtifact.major_font_family_east_asia` wire field as the contract and verify generated bindings remain unchanged

## 2. Codec

- [x] 2.1 Project the observed major East Asian typeface from the canonical shared ThemePart with a hash-bound script-specific capability
- [x] 2.2 Validate one-field source-bound font-scheme ownership and patch only `a:majorFont/a:ea/@typeface` while preserving the rest of the package

## 3. Regression

- [x] 3.1 Add a focused authored → source projection → major-East-Asian-font edit → second projection test covering no-op bytes, capability authority, changed-part scope, non-target package preservation, and Open XML validity
- [x] 3.2 Run the focused native test, schema/protocol checks, and presentation skill maintenance check
