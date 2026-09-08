# Presentation capability development guide

This guide applies to `skills/presentations/` and presentation-facing changes
that it documents. The implementation sources remain `src/ppj/`,
`native/OfficeKit/src/OfficeKit.Codec/`, and the presentation test/eval trees.

## A capability is complete only when all surfaces agree

For every new or changed presentation capability, trace the feature through
these surfaces before calling it complete:

1. PPJ schema and semantic validation.
2. C# authored/imported compiler and the relevant native-ref or opaque boundary.
3. Public help and generated PPJ reference documentation.
4. Agent guidance and the presentation Skill route.
5. OfficeKit structural inspection and review evidence.
6. Visual preview coverage, or an explicit partial/opaque diagnostic.
7. PPTX export, re-import, and source/edit fidelity where the feature is
   source-bound.
8. Focused fixtures and regression tests.

`src/ppj/capability-registry.json` is the discoverability ledger. Do not create
a second untracked feature list in a renderer or Skill. Add the owner,
boundary, support level, and test/fixture reference to the registry instead.

## Visual preview contract

The local preview renderer is a review aid, not a promise of PowerPoint
compatibility. It must consume the same validated PPJ meaning as the compiler,
preserve element IDs and z-order, and report unsupported content explicitly.

- `supported`: render the complete declared visual state.
- `partial`: render the supported projection and emit machine-readable limits.
- `opaque`: retain a source preview or bounded placeholder; never flatten or
  invent semantics.
- `unavailable`: report the missing dependency or input; never claim a pass.

Every new visual field must therefore have either a renderer mapping and visual
fixture, or an explicit fallback/diagnostic test. A feature may not silently
export successfully while preview drops it.

## Reliability invariants

- Missing data stays missing; do not connect through an unknown observation or
  coerce it to zero.
- Connector geometry and chart marks must follow the declared data/topology
  graph, not visual proximity.
- Source-bound and opaque payloads remain source-bound and opaque.
- Unknown image alpha, subject bounds, crop, edge, or shadow information is
  handled conservatively.
- Check, build, render, review, and (for imported content) re-import are one
  evidence chain. A later stage cannot be replaced by an earlier pass.

## Change workflow

Before implementation, identify the owning PPJ path and capability-registry
entry. During implementation, update the narrowest affected schema/compiler,
renderer projection, Skill/reference, and fixture together. Before delivery,
run the focused tests first, then the presentation portability/reference-sync
gates and the applicable OfficeKit build checks. Record environment-dependent
renderer skips as evidence; do not convert them into visual success.

When a capability is intentionally unsupported, document the boundary and add
one negative fixture. When it is partial, document which fields are preserved,
which are approximated, and how a human should review the result.

