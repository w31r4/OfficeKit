# Trendline list evidence

## Contract

The existing `data.series[].trendlines` array now supports source-bound
insertion, reordering, replacement, removal and clearing on recognized
bar/column and line series, including categorical combo series. Omission and
an empty array both clear the list. Schema and wire shape are unchanged.

## Focused experiment

Run from `/tmp` using SDK 8.0.417, leaving the pinned repository SDK unchanged:

```sh
DOTNET_ROOT=/home/zenfun/.dotnet DOTNET_CLI_HOME=/tmp/officekit-dotnet-cli NUGET_PACKAGES=/home/zenfun/.nuget/packages /home/zenfun/.dotnet/dotnet test /home/zenfun/mywork/OfficeKit/native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj --filter 'FullyQualifiedName~PpjTrendlineList|FullyQualifiedName~OpenXmlChartTrendlineCodecTests|FullyQualifiedName~OpenXmlChartErrorBarsCodecTests' --no-restore -p:WarningLevel=0
```

Passed 11, failed 0, skipped 0.

- Four lifecycle cases: column, horizontal bar, line, and combo secondary line.
  Each starts with two authored trendlines, inserts in the middle, reorders
  and replaces entries, shrinks, omits, adds from empty, clears with `[]`, then
  recreates a styled trendline.
- Every step removes embedded PPJ before projection, checks native list order
  and values, validates the native series with the Office 2021 Open XML schema,
  and checks unchanged ChartML outside the target list, including error bars
  and sibling chart data. All other ZIP entries retain their bytes.
- Two ordinary/combo malformed-owner cases retain source ChartPart bytes and
  withhold analytics editing for labels, extensions, duplicate type nodes and
  unknown types.
- Invalid polynomial state, moving-average period, fractional forecast and
  more than sixteen entries are rejected when inserted.
- Shared trendline regression now checks list shrinkage and rejects deletion
  of an unsupported native owner without altering it. Existing error-bar
  regressions pass.

The XLSX adapter-level count restriction in XlsxChartCodec and its JavaScript
adapter remains unchanged. This change adds no XLSX list-editing capability.

These experiments establish native structure and source-bound mutation;
PowerPoint host rendering was not part of this field change.

## Repository checks

- Strict OpenSpec validation: passed.
- Presentation Skill maintainer sync/check: passed.
- Capability matrix regenerated; reference Skill source sync: passed.
- `git diff --check`: passed.
- Skill portability: the existing `installed REPL guidance: powerpoint-live-control`
  assertion still fails. This test and the PowerPoint Live Skill are unchanged
  by the trendline work. No full-suite pass is claimed.
- Protocol/schema shape unchanged; protobuf regeneration was not required.
