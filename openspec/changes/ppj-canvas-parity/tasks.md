## 1. Language and projection

- [x] 1.1 Add canvas nativeRef plus `setCanvas` vocabulary.
- [x] 1.2 Project an exact hash-bound canvas capability.

## 2. Source-bound lowering

- [x] 2.1 Convert capable PPJ point dimensions to deterministic EMUs.
- [x] 2.2 Reuse the existing native canvas-only writer and report affected pages.

## 3. Agent surface and lean verification

- [x] 3.1 Regenerate `ppj.md` and update review guidance and coverage.
- [x] 3.2 Extend the existing comprehensive PPJ contract and second projection.
- [x] 3.3 Run focused C# build/test, Skill-maintainer, and strict OpenSpec checks.
- [ ] 3.4 Commit atomically and fast-forward main without force pushing.

## Evidence

- `dotnet build native/OfficeKit/src/OfficeKit.Codec/OfficeKit.Codec.csproj --no-restore`
  succeeded with zero warnings and errors.
- The existing comprehensive
  `PpjV1CompilesCanonicalPresentationProgramDeterministically` contract passed
  once after combining one canvas-width edit with the existing three-page
  reorder. It proved that only `ppt/presentation.xml` changed, every page was
  reported affected, and second projection recovered the requested width,
  unchanged height, stable page IDs, stable page-local element IDs, comments,
  sections, and custom shows.
- The Presentation Skill maintainer passed with 151 Help APIs, 73 native leaves,
  and 13 host-only operations.
- `npx openspec validate ppj-canvas-parity --strict` passed.
- No new test file, fixture, protocol field, full suite, or package gate was
  added or run.
