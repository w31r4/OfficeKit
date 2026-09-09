## 1. Field and native effects

- [x] 1.1 Add schema/wire and shared direct soft-edge value checks, extending the chart effects adapter; verify native ordering, malformed graphs and explicit zero/wire presence.
- [x] 1.2 Connect authored/source-bound/projector/vector and style-precedence paths, including the discovered direct-effect theme-reference validation mismatch; verify line/combo lifecycle, independent sibling deletion, fresh projection and non-target ZIP preservation.

## 2. Documentation and evidence

- [x] 2.1 Update backlog, registry, chart reference and coverage; regenerate references and pass Skill/capability checks.
- [x] 2.2 Run focused native effects regressions, JS preservation, proto check and strict OpenSpec; record actual evidence and leave full F-07 open.

## Evidence

Implemented from `5837c51e` in an isolated worktree. Chart text softEdge uses additive wire field 21, the ordinary required radius schema, shared direct-value validation, and native order glow → outerShdw → softEdge. Ordinary Presentation effect composition remains in its existing codec. The leaf is parsed through a composite effect list after checking it has no extra children.

SDK 8.0.128 final focused run passed 55/55, zero skips: `PpjChartTextSoftEdge`, `PpjChartTextGlow`, `PpjChartTextShadow`, `PpjChartTextHighlight`, `PpjTrendlineRichText`, `PpjTrendlineLabel`, and existing `SoftEdge`/`Glow`/`Shadow` regressions. Four new cases cover line/combo independent three-effect lifecycle, native ordering/bounds/precision, source-bound fresh projection and non-target ZIP preservation, plus field precedence/vector zero overrides. Eight malformed soft-edge graphs remain guarded; the earlier valid glow-plus-softEdge negative fixture now uses an oversized radius.

The initial authored theme-colored sibling experiment exposed generic resource validation rejecting standard effect theme tokens before compilation. The validator now permits recognized undeclared theme colors only for shadow/glow effect colors. Authored/source-bound lifecycle preserves accent1/accent2 identity; explicit wrong-kind grammar tokens and unsupported foreground theme colors remain rejected by focused assertions. This is a fix to the existing effect contract, not a new color fallback for all fields.

JS worksheet chart preservation, `npm run proto:check`, Skill maintenance, generated capability matrix check, preview capability coverage, Skill portability, strict OpenSpec and whitespace checks passed. Generated references use the checked-in scripts. No NativeAOT rebuild, renderer or PowerPoint host acceptance is claimed; full F-07 remains open.
