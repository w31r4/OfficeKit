# logBase evidence

## Implementation

`chartAxis.logBase` and radar `spokeAxis.logBase` share optional wire field 25
and native `c:scaling/c:logBase`. The [SDK ChartML schema](https://github.com/dotnet/Open-XML-SDK/blob/main/data/schemas/schemas_openxmlformats_org_drawingml_2006_chart.json)
specifies a double from 2 through 1000 inclusive, preceding orientation.
Deletion removes the native node. No wire-version bump or new dependency.

## Focused experiments

SDK 8.0.417, run from `/tmp` with the absolute codec test project path to leave
the repository's pinned `global.json` unchanged:

```sh
DOTNET_ROOT=/home/zenfun/.dotnet DOTNET_CLI_HOME=/tmp/officekit-dotnet-cli NUGET_PACKAGES=/home/zenfun/.nuget/packages /home/zenfun/.dotnet/dotnet test /home/zenfun/mywork/OfficeKit/native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj --filter 'FullyQualifiedName~PpjAxisLogBase|FullyQualifiedName~PpjValueAxisMinorUnit|FullyQualifiedName~PpjChartRoundedCorners|FullyQualifiedName~PpjChartAxisPositions' --no-restore -p:WarningLevel=0
```

Passed: 16, failed: 0, skipped: 0. Ten new logBase cases cover:

- Column value axis, scatter/bubble numeric x axis, combo primary/secondary
  value axes, and both radar generic/spoke authoring forms.
- Token-authored 10 → token edit 2.5 → remove → add 2 → edit 1000 → remove.
  Every projection strips embedded PPJ. Assertions inspect actual native
  values, ordering, axis identity, unchanged remaining ChartML and unchanged
  bytes for every unrelated ZIP entry.
- Invalid literals, token values/kinds, category-axis use and nonpositive
  explicit limits fail both authored and source-bound compilation.
- Ordinary/combo malformed native nodes with bad values, duplicate owners,
  missing value, extra attributes or children withhold axis editing; no-op
  compilation preserves the original ChartPart.

The radar experiment exposed a projection-route gap after deletion; adding
the same field to spokeAxis closed repeated add/remove behavior.

`node test/worksheet-chart-axis-preservation.mjs`: passed. This small regression,
also imported by the normal spreadsheet test, exercises wire → JavaScript
chart → title edit → wire and checks that the original logBase survives. It
protects XLSX users from a field loss introduced by the newly editable shared
native axis profile.

## Repository checks

- `npm run proto:check`: passed after staging regenerated bindings.
- `openspec validate ppj-chart-axis-log-base --strict`: passed.
- Presentation Skill maintainer `sync` then `check`: passed.
- `node test/reference-skill-sync.mjs`: passed.
- `git diff --check`: passed.
- `node test/skill-portability.mjs`: fails on the existing PowerPoint Live
  REPL-guidance assertion. Both the test and that Skill are unchanged from HEAD.
- `node test/reference-skills.mjs`: cannot complete because `pdftoppm` is absent.
- Five existing XLSX chart regressions: two passed; three failed. Re-running
  the same three on independent pre-change checkout `2f338d23` reproduced
  identical failures: area stacking read-only expectation at line 4120,
  bubbleScale read-only expectation at line 4331, and legend-position
  validation at line 3693 of XlsxCodecTests.cs. These failures predate logBase.

This evidence covers native structure and source-bound round trips. It does
not claim PowerPoint host rendering or a green full test suite.
