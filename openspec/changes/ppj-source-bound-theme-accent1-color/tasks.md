## 1. Contract

- [x] 1.1 Extend the PPJ theme capability vocabulary and source-bound coverage entry for `design.theme.accentColors.accent1`, then validate the OpenSpec change
- [x] 1.2 Reuse the existing `PresentationThemeArtifact.accent_rgb` wire field and verify generated bindings remain unchanged

## 2. Codec

- [x] 2.1 Read six strict direct RGB accents from one canonical shared ThemePart and project the observed accent palette with a hash-bound accent1 capability
- [x] 2.2 Validate source-bound accent ownership and patch only `a:accent1Color/a:srgbClr/@val` while preserving the rest of the package

## 3. Regression

- [x] 3.1 Add a focused authored → source projection → accent1 edit → second projection test covering no-op bytes, capability authority, changed-part scope, non-target package preservation, and Open XML validity
- [x] 3.2 Run the focused native test, schema/protocol checks, and presentation skill maintenance check
