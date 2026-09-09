## 1. Position lifecycle

- [x] 1.1 Extend chart schema, wire, direct reader/writer and projection with optional position state; verify JS/native presence, bounds and ordinary canonical/full-span guard preservation.
- [x] 1.2 Extend the line/combo reflection regression with independent start/end edits and removals, equal/reversed positions, precision, named precedence and vector defaults; verify source no-op, fresh projection and non-target ZIP preservation.

## 2. References and evidence

- [x] 2.1 Update chart reference, registry, coverage/backlog and generated references; verify maintenance, capability and portability checks.
- [x] 2.2 Run focused reflection/effect tests, proto and strict OpenSpec checks; record actual results and remaining boundaries before scoped atomic publication.

Evidence (2026-09-09, base 0373a7cb): final focused native reflection/effect suite passed 69/69, zero skipped, using SDK 8.0.128. Existing chart reflection cases now cover absent positions, independent end edits/start removal, explicit 0/1, equal/reversed values, precision, named reflection precedence and vector default/run behavior across line/combo lifecycles. Native XML, fresh projections, exact source no-op, other reflection/effect values, literal text and non-target ZIP entries are asserted. Invalid position/token/null/unknown-child cases remain rejected. Ordinary reflection tests retain full-span native output, partial-span source ownership and the fixed fractional-angle native leaf behavior.

JS worksheet-chart retention and optional wire fields 6/7 passed. Proto lint/generation/drift, Skill maintenance, portability (255 files), root reference sync (333 files), capability generation/coverage, strict OpenSpec and whitespace checks passed. Native defaults were checked against the upstream schema/getters referenced in design.md; their presence remains unsynthesized on chart import/write. No NativeAOT package rebuild, host visual acceptance or complete F-07 claim.
