## 1. Persistent label state

- [x] 1.1 Add the label schema/wire contract and shared native codec, regenerate bindings and verify protocol checks plus native label structure.
- [x] 1.2 Connect authored/source-bound grammar and projection; verify ordinary/combo label creation, replacement, empty/default presence, removal and recreation with a minimal original-source/fresh-projection test.
- [x] 1.3 Verify malformed/unsupported label owners remain source-owned and analytics mutations emit no file; rerun neighboring trendline and error-bar tests.

## 2. Discoverability and evidence

- [x] 2.1 Update the chart reference, registry and F-07; regenerate/check derived documentation, validate OpenSpec strictly and record exact focused results and remaining visual/layout limits.

Evidence (2026-09-09): the managed `PpjTrendlineLabel`, `PpjTrendlineList`,
`OpenXmlChartTrendlineCodecTests` and `PpjErrorWorkbook` filter passed 24/24.
The label experiment checks literal and grammar-token text/format, all five
properties, default-object presence, deletion/recreation, native node validation,
unchanged unrelated ChartML/ZIP entries and fresh source projection. Rejection
cases cover invalid values, formula text, unmodeled text, manual layout,
source-linked formats, duplicate nodes and extensions.

`node test/worksheet-chart-axis-preservation.mjs` passes with an added imported
trendline-label preservation assertion for an unrelated JS chart title edit.
`npm run proto:check` passes against the isolated change snapshot after staging
the regenerated binding; regeneration is deterministic. Presentation Skill
maintenance passes (151 Help APIs, 278 native leaves), reference Skill sync
passes (333 files), and strict OpenSpec validation and diff checks pass.

The shared worktree also passes skill-portability (255 files), with a concurrent
change to that test's PowerPoint Live expectation. That test adjustment belongs
to the separate preview work and is not included in this field commit.
This evidence does not establish host visuals, automatic label layout or local
SVG label rendering. Complex labels and the complete F-07/P0/P1 backlog remain
open. Generated docs use an isolated registry to keep the concurrent preview
changes separate from this atomic field increment.
