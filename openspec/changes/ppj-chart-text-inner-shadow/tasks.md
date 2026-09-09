## 1. Field lifecycle

- [x] 1.1 Add schema/wire and shared native effect parsing/writing; verify optional presence, supported ordering, bounds and source-owned invalid graphs with focused native/JS regressions and proto check.
- [x] 1.2 Connect authored/source-bound compilation, projection, precedence and vector defaults; verify line/combo creation/change/deletion/recreation, exact no-op and unrelated ZIP preservation plus explicit vector overrides.

## 2. Discoverability and evidence

- [x] 2.1 Update capability registry, focused chart reference, coverage/backlog and generated references; verify maintenance, portability and capability checks.
- [x] 2.2 Run focused effect regression suite and strict OpenSpec validation; record actual results and remaining boundaries before atomic publication.

Implementation evidence (2026-09-09, base d891a173): the focused native filter covering inner/outer shadows, glow, soft edge, chart highlight and trendline rich text/labels passed 59/59 with zero skips using SDK 8.0.128. The new line/combo lifecycle checks theme identity, omitted geometry/alpha versus explicit zero, color/opacity grammar resolution, clearing/deletion/recreation, retained glow/outer-shadow/soft-edge order, literal text, exact source no-op and every non-target ZIP entry across fresh projections. Native-only checks prove inner-shadow-only styles, full effect order under OpenXmlValidator, bounds, tokens/precision, named field precedence and vector overrides; ten malformed imported inner-shadow graphs remain source-owned.

JS worksheet-chart preservation passed with optional wire field 22 and BigInt angle/geometry values. An initial test fixture used Number zero for the existing int64 angle; it was corrected to BigInt without changing the wire contract. Proto lint/generation/drift check, Skill maintenance, portability (255 files), capability generation/coverage, strict OpenSpec and whitespace checks passed. No NativeAOT package rebuild, preview/PowerPoint appearance or whole F-07 completion is claimed. The separately published preview test failures are outside this field's code and validation scope.
