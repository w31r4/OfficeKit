## 1. Contract and native implementation

- [x] 1.1 Extend chart shadow schema and additive wire fields; regenerate bindings and pass proto and JS wire preservation checks.
- [x] 1.2 Propagate transforms through chart compilation/projection and shared native writing with a chart-only read opt-in; verify native bounds and ordinary imported guards.

## 2. Minimal experiment and documentation

- [x] 2.1 Extend existing line/combo shadow lifecycle, style/vector and precision tests; pass focused effect tests with independent siblings, fresh projection, no-op and non-target ZIP assertions.
- [x] 2.2 Update registry, focused reference, coverage, backlog and generated docs; pass maintenance, capability, portability, reference sync and strict OpenSpec checks with exact verification scope recorded.

Verification: SDK 8.0.128 focused effect suite passed 69/69, zero skips. The first run passed 67 and failed both line/combo lifecycle cases because glow writing rejected a transformed shadow sibling. The final implementation adds a guarded write opt-in for chart reconstruction and just-rebuilt run/default-run shadows; unchanged ordinary read/write guards are asserted, and vector shadow-plus-glow cases now pass. Existing lifecycle checks retain all four sibling effects, literal text, non-target ZIP entries, exact fresh no-op and reprojection. Native bounds/precision and optional attributes, named precedence, vector overrides and JS wire retention pass. Proto, maintenance, capability matrix/coverage, portability and strict OpenSpec checks pass. No full repository test, NativeAOT rebuild or host display acceptance was performed.
