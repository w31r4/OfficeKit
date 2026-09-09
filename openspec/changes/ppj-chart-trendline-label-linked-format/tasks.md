## 1. Native flag state

- [x] 1.1 Add schema field/dependency and compatible wire enum; regenerate bindings and verify `npm run proto:check` plus JS enum round-trip/default preservation.
- [x] 1.2 Implement shared label read/write/validation and both PPJ compiler paths/projector; verify line/combo authored and source-bound true/false/null lifecycle, removal/recreation, native format attributes and non-target bytes.

## 2. Rejection and field documentation

- [x] 2.1 Replace the newly supported native fixture with an invalid flag and add missing-format/type/wire rejection checks; verify no-op preservation and no output for rejected edits in focused native regressions.
- [x] 2.2 Update registry, backlog and chart reference with exact host boundary; regenerate manual/matrix and verify maintainer, generators, Skill portability/reference sync, OpenSpec strict validation and diff check.

Evidence (2026-09-09): isolated snapshot passes 18/18 managed tests filtered by `PpjTrendlineLabel|PpjTrendlineList|OpenXmlChartTrendlineCodecTests`, including exact native attributes, token formats, fresh projection, no-op and other ZIP bytes, format removal/recreation and null-state retention during a label text edit. `test/worksheet-chart-axis-preservation.mjs`, `npm run proto:check`, matrix generation/check, preview capability coverage, presentation maintainer (151 Help APIs / 278 native leaves), Skill portability (255 files), strict OpenSpec and diff checks pass. Reference sync passes against the main workspace's local reference source (333 files). Test log: `/tmp/officekit-trendline-linked-format-test.log`.

This is native flag state evidence only. It does not establish workbook format synchronization, NativeAOT release or host/SVG rendering. Formula/rich-text labels, complex effects/extensions and the other F-07/P0/P1 gaps remain open.
