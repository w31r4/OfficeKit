# Error-bar object lifecycle evidence

Verified on 2026-09-09. The existing PPJ `data.series[].errorBars` object now supports source-bound addition, modification, omission/removal and recreation for recognized column, bar, line and categorical combo series. No schema or wire fields were added.

## Focused native regression

From `/tmp`, using the installed .NET 8 SDK to avoid the repository's unavailable pinned SDK:

```sh
env DOTNET_ROOT=/home/zenfun/.dotnet DOTNET_CLI_HOME=/tmp/officekit-dotnet-cli NUGET_PACKAGES=/home/zenfun/.nuget/packages \
  /home/zenfun/.dotnet/dotnet test \
  /home/zenfun/mywork/OfficeKit/native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj \
  --filter 'FullyQualifiedName~PpjErrorBarsLifecycle|FullyQualifiedName~PpjTrendlineList|FullyQualifiedName~OpenXmlChartErrorBarsCodecTests|FullyQualifiedName~OpenXmlChartTrendlineCodecTests' \
  --no-restore -p:WarningLevel=0
```

Result: **22 passed, 0 failed, 0 skipped**. The managed shared, presentation, spreadsheet and test assemblies were rebuilt.

- Four chart profiles repeatedly remove, insert zero, switch scalar modes/direction/type/cap/stroke, remove again and recreate. Native series pass Office 2021 Open XML validation. Fresh projection uses PPTX bytes with the private PPJ snapshot removed.
- Every edit changes only the selected errorBars owner in ChartML; trendlines, the other series, axes and data stay equal, and every other ZIP entry stays byte-identical.
- Duplicate owners/directions, extensions and invalid scalar values retain their original ChartPart without analytics edit capability. Shared native deletion rejects malformed/custom owners before modifying XML.
- Literal custom plus/minus sources survive no-op compilation; an attempted scalar insertion is rejected because the native custom owner is absent from the PPJ projection.
- Invalid inserted scalars, missing required values, standard-error with a scalar, unknown custom mode and invalid direction are rejected. Adjacent trendline and shared-owner regressions pass.

## Documentation and preview checks

These commands passed from the repository root:

```sh
node skills/presentations/skills/presentation-skill-maintainer/scripts/maintain-presentation-skill.mjs sync
npm run docs:presentation-capabilities
node skills/presentations/skills/presentation-skill-maintainer/scripts/maintain-presentation-skill.mjs check
node test/ppj-preview-output-evidence.mjs
node test/ppj-svg-preview.mjs
node test/reference-skill-sync.mjs
openspec validate ppj-chart-error-bars-lifecycle --strict
git diff --check
```

The dependency-light preview regression checks that adding errorBars to each supported profile emits `partial` / `chart-error-bars-not-rendered`, and removing the object removes that diagnostic. The real codec/sharp smoke also passes; it verifies output publication, not error-bar geometry.

`node test/skill-portability.mjs` was run and fails at the existing `installed REPL guidance: powerpoint-live-control` assertion: it expects `officekit repl` or `references/repl.md`, while the Live Skill documents typed `officekit live` operations. Neither that test nor the Live Skill is changed here.

## Remaining boundaries

Custom plus/minus PPJ data, formula/workbook synchronization, trendline labels/effects and broader F-07 topology remain backlog work. XLSX adapter presence policy is unchanged. The local SVG does not draw error bars and now reports that limit explicitly. No host-PowerPoint, visual-fidelity, full-suite or NativeAOT release acceptance is claimed.
