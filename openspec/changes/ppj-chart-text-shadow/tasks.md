## 1. Field and native semantics

- [x] 1.1 Add chart shadow schema/wire and strict shared native effects support; verify optional presence, invalid graphs, native XML ordering and protobuf preservation.
- [x] 1.2 Connect authored, source-bound, projector and vector/style-precedence paths; verify line/combo lifecycle, tokens, explicit zero/false, deletion/recreation and non-target ZIP preservation.

## 2. Evidence and documentation

- [x] 2.1 Update backlog, chart reference, coverage and registry/generated references; verify documentation generation and capability checks.
- [x] 2.2 Run focused native/JS regressions, proto check and strict OpenSpec validation; record actual results and retain full F-07 as partial.

## Evidence

Implemented from main `8b5b6771` in an isolated worktree. `chartTextStyle.shadow` uses additive wire field 19 and shared direct outer-shadow validation; optional geometry/alpha/rotation are preserved by the dedicated chart mapping. The native schema validator only resolves top-level definitions, so the initial nested-property references were replaced with equivalent local property definitions before the passing run.

SDK 8.0.128: final focused run passed 35/35, zero skips, after aligning wire geometry validation with PPJ ranges. Selection: `PpjChartTextShadow`, `PpjChartTextHighlight`, `PpjTrendlineRichText`, `PpjTrendlineLabel`, and existing `Shadow` tests. The four new cases cover line/combo lifecycle, native ordering and bounds/tokens, plus field precedence/vector overrides; the existing label guard also covers 11 unsupported effect graphs. No-op preserves original bytes; each changed chart preserves all non-target ZIP entries, literal text, theme identity and explicit zero/false versus omitted fields across fresh projection.

`node test/worksheet-chart-axis-preservation.mjs`, `npm run proto:check`, presentation Skill maintenance check, generated capability matrix check, PPJ preview capability coverage and Skill portability passed. Generated references were synchronized using the checked-in scripts. Strict OpenSpec and whitespace checks passed. These are native structure/round-trip and discoverability checks; no NativeAOT rebuild, renderer/PowerPoint host acceptance or full F-07 completion is claimed.
