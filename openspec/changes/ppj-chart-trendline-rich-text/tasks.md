## 1. Structured field and native codec

- [x] 1.1 Add structured label.text schema and additive rich-text wire messages; regenerate bindings and verify schema invalid cases, JS preservation and `npm run proto:check`.
- [x] 1.2 Implement strict shared rich-text codec and both compiler/projector paths using chart style helpers; verify ordinary/combo authored/source-bound lifecycle, style/token/whitespace/break preservation, simple-text canonicalization and exact non-target bytes with focused native tests.

## 2. Boundaries and discoverability

- [x] 2.1 Add unsupported native graph, invalid run/style/dual-wire-state and resource-budget checks; verify no-op input preservation and rejection without output.
- [x] 2.2 Update chart reference, backlog and registry; regenerate manual/matrix and verify maintainer, generated coverage, Skill portability/reference sync, OpenSpec strict validation and diff check. Record bounded evidence and remaining text graphs.

Evidence (2026-09-09): the final isolated snapshot passes 21/21 managed tests filtered by `PpjTrendlineRichText|PpjTrendlineLabel|PpjTrendlineList|OpenXmlChartTrendlineCodecTests`, including line/combo lifecycle, empty text tokens, paragraph/run/end styles, explicit false, whitespace, breaks, canonical simple text, non-target ZIP bytes, malformed graphs and budgets. The first fixture used the wrong size-token kind; fixing it exposed the shared string resolver's empty-run rejection, which was fixed with the scoped allowEmpty option. Final log: `/tmp/officekit-trendline-rich-text-final-test.log`.

`npm run proto:check`, JS chart/label wire preservation, both capability generator checks, preview capability coverage, presentation maintainer (151 Help APIs / 278 native leaves), Skill portability (255 files), strict OpenSpec and diff checks pass in the isolated snapshot. Reference sync passes against the main workspace's local reference source (333 files). Concurrent preview entrypoint code is excluded from the isolated compiler snapshot and commit.

Formula references, fields, hyperlinks, nonempty body/list formatting, unsupported character properties/effects and other F-07/P0/P1 gaps remain open. These managed/native tests do not establish rebuilt NativeAOT release or host/SVG typography fidelity.
