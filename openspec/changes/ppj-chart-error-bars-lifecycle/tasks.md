## 1. Scalar object lifecycle

- [x] 1.1 Implement optional-object edits across PPJ compilation, PPTX topology checks and native owner patching; verify add/edit/delete/recreate for column, bar, line and combo through `PpjErrorBarsLifecycle` regressions.
- [x] 1.2 Protect malformed owners and unprojected custom data; verify unchanged native XML on rejected deletion, no-op preservation, custom insertion rejection and invalid scalar insertions in focused native/PPJ tests.

## 2. Documentation and adjacent verification

- [x] 2.1 Update F-07, capability registry and chart reference, regenerate PPJ reference and capability matrix, and verify maintenance checks. Keep custom-data limitations explicit and verify a partial preview diagnostic for errorBars with `test/ppj-preview-output-evidence.mjs`.
- [x] 2.2 Run adjacent trendline/shared error-bar tests, strict OpenSpec validation and diff checks; record commands, results and validation limits in evidence.md.
