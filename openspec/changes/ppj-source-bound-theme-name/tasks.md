## 1. Contract

- [x] 1.1 Add the source-bound theme-name requirement and validate the OpenSpec change
- [x] 1.2 Extend the PPJ theme schema and capability/coverage documentation for `design.theme.name`

## 2. Codec

- [x] 2.1 Read a unique existing ThemePart name during PPTX import and project its stable theme native reference
- [x] 2.2 Validate the theme name owner in source-bound compilation and patch only `a:theme/@name` during export

## 3. Regression

- [x] 3.1 Add a focused authored → source projection → name edit → second projection test covering changed-part scope, no-op bytes, authority, and Open XML validity
- [x] 3.2 Run the focused native test, schema/protocol checks, and presentation skill maintenance check
