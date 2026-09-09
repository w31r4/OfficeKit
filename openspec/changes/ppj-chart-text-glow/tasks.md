## 1. Field and effect composition

- [x] 1.1 Add schema/wire, shared glow ownership and strict chart effects composition; verify native order, invalid graphs, meaningful style and wire retention.
- [x] 1.2 Connect authored/source-bound/projector, shared color/opacity resolution and vector precedence; verify line/combo lifecycle with independent sibling deletion and fresh projection.

## 2. Documentation and evidence

- [x] 2.1 Update backlog, registry, chart reference and coverage; regenerate references and pass capability/Skill checks.
- [x] 2.2 Run focused native glow/shadow/chart regressions, JS preservation, proto check and strict OpenSpec; record actual evidence without closing the full F-07.

## Evidence

Implemented from `90202df2` in an isolated worktree. The shared chart text effects adapter accepts one glow, one outer shadow, or glow followed by outer shadow. All other effect graphs retain source ownership. Chart glow uses additive wire field 20 and the existing required-color/radius schema. The shared chart color/opacity resolver retains theme identity, declared grammar values and alpha presence; ordinary text mapping remains unchanged.

SDK 8.0.128 focused run passed 45/45, zero skips: `PpjChartTextGlow`, `PpjChartTextShadow`, `PpjChartTextHighlight`, `PpjTrendlineRichText`, `PpjTrendlineLabel`, and existing `Glow`/`Shadow` tests. Four new cases cover line/combo lifecycle with independent sibling removal/restoration, original-byte no-op, fresh projection, native effect ordering, non-target ZIP bytes, tokens/bounds and vector field precedence. The existing negative fixture was corrected to use oversized glow now that valid glow is supported; seven further malformed effect graphs remain guarded.

JS worksheet chart preservation, `npm run proto:check`, Skill maintenance, generated capability matrix check, PPJ preview capability coverage, Skill portability, strict OpenSpec and whitespace checks passed. An initial proto insertion matched a second message; it was corrected before successful generation/build. Generated files use the checked-in scripts. No NativeAOT rebuild, renderer or PowerPoint host acceptance is claimed; native glow without a radius and the full F-07 remain open.
