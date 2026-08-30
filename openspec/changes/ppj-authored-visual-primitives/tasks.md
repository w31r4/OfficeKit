## 1. Truthful compiler surface

- [x] 1.1 Inventory present PPJ visual fields that are ignored by the authored compiler
- [x] 1.2 Reject every unsupported present visual field before output
- [x] 1.3 Generate authored support and boundary metadata from the registry

## 2. Geometry

- [x] 2.1 Lower custom PPJ paths into the existing native custom-geometry IR
- [x] 2.2 Preserve deterministic viewBox scaling and validate path budgets
- [x] 2.3 Keep preset adjustments explicit and fail-closed

## 3. Paint and layers

- [x] 3.1 Add bounded linear/radial gradient state and native shape/background writers
- [x] 3.2 Add bounded shape/connector line opacity
- [x] 3.3 Keep imported unknown paint graphs source-preserved

## 4. Data visuals

- [x] 4.1 Compile PPJ legend, stacking, gap, axis, gridline, chart-area, and plot-area style
- [x] 4.2 Compile PPJ table cell fill, text style, and borders where topology is supported
- [x] 4.3 Reject still-unsupported chart/table properties explicitly

## 5. Agent surface and delivery

- [x] 5.1 Regenerate `ppj.md` and focused shapes/charts/tables guidance
- [x] 5.2 Extend the existing integrated PPJ test without creating a new matrix
- [x] 5.3 Run narrow native, proto, generated-reference, and OpenSpec checks
- [x] 5.4 Commit atomically, push the feature branch, and fast-forward main
