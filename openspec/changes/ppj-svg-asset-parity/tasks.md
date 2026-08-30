## 1. Language and projection

- [x] 1.1 Add source-bound `image.svgAsset` and `replaceSvg` vocabulary.
- [x] 1.2 Project both members of a proven raster/SVG fallback pair.

## 2. Source-bound lowering

- [x] 2.1 Validate the replacement SVG asset MIME and hash.
- [x] 2.2 Reuse the native picture writer without changing fallback topology.

## 3. Agent surface and lean verification

- [x] 3.1 Regenerate `ppj.md` and update media/layer review guidance and coverage.
- [x] 3.2 Extend one existing comprehensive PPJ contract with a paired SVG.
- [x] 3.3 Run focused C# build/test, Skill-maintainer, and strict OpenSpec checks.
- [x] 3.4 Commit atomically and fast-forward main without force pushing.

## Evidence

- The existing comprehensive
  `PpjV1CompilesCanonicalPresentationProgramDeterministically` contract passed
  once after isolating one real PNG + `asvg:svgBlip` pair. It proved
  byte-identical no-op, `replaceSvg/image.svgAsset`, unchanged fallback hash,
  stable image ID, replacement SVG hash recovery, and no change to unrelated
  SlideParts.
- The same focused `dotnet test` invocation rebuilt the codec and test project
  successfully before running the single filtered contract.
- The Presentation Skill maintainer passed with 151 Help APIs, 73 native
  leaves, and 13 host-only operations after regenerating `ppj.md`.
- `openspec validate ppj-svg-asset-parity --strict` passed.
- No new test file, protocol field, full suite, package gate, SVG node DSL, or
  fallback rasterizer was added or run.
