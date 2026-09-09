## 1. Field and native lifecycle

- [x] 1.1 Add schema and additive protobuf layout/manual fields and regenerate bindings; verify schema rejection and `npm run proto:check`.
- [x] 1.2 Implement shared layout validation/read/write/PPJ mapping in both compiler routes and projector; verify line/combo authored and source-bound layout lifecycle, container/default/zero presence, native XML validation and non-target byte preservation with focused native tests.

## 2. Boundaries and discoverability

- [x] 2.1 Update unsupported fixtures and add invalid/opaque layout coverage; verify no-op bytes, blocked analytics and no output for invalid edits, plus JS wire preservation.
- [x] 2.2 Update backlog, registry and chart reference, regenerate manual/matrix; verify generated checks, Skill portability/reference sync, strict OpenSpec validation and diff check. Record actual test evidence and remaining F-07 gaps.

Evidence (2026-09-09): the isolated publication snapshot passes 15/15 managed tests selected by `PpjTrendlineLabel|PpjTrendlineList|OpenXmlChartTrendlineCodecTests`; the same tests passed in the shared workspace. `test/worksheet-chart-axis-preservation.mjs` passes, including protobuf preservation of empty containers and optional zero. `npm run proto:check`, both capability generator checks, preview capability coverage, presentation maintainer (151 Help APIs / 278 native leaves), Skill portability (255 files), strict OpenSpec and diff checks pass. Reference sync passes in the main workspace (333 files); its first isolated run lacked the ignored legacy reference source, so it was checked in the workspace containing that source. Native test log: `/tmp/officekit-trendline-layout-isolated-test.log`.

This closes the bounded trendline label layout increment. Formula/rich-text labels, source-linked formats, layout extensions, complex effects and other F-07 topology/data gaps remain. Managed/native structure and round-trip tests do not establish NativeAOT release or host/SVG positioning fidelity.
