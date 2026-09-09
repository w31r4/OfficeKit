# Literal custom error data evidence

Verified on 2026-09-09. PPJ now owns `errorBars.valueType: "custom"` and literal `plus/minus: { values, formatCode? }`. Arrays align with every series point; zero, asymmetry, side selection and optional format are preserved. Existing protobuf fields are reused without a wire-version change.

## Native regression

Run from `/tmp` with installed .NET 8, avoiding the unavailable repository-pinned SDK:

```sh
env DOTNET_ROOT=/home/zenfun/.dotnet DOTNET_CLI_HOME=/tmp/officekit-dotnet-cli NUGET_PACKAGES=/home/zenfun/.nuget/packages \
  /home/zenfun/.dotnet/dotnet test \
  /home/zenfun/mywork/OfficeKit/native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj \
  --filter 'FullyQualifiedName~PpjCustomErrorData|FullyQualifiedName~PpjErrorBarsLifecycle|FullyQualifiedName~OpenXmlChartErrorBarsCodecTests|FullyQualifiedName~PpjTrendlineList' \
  --no-restore -p:WarningLevel=0
```

Result: **26 passed, 0 failed, 0 skipped** after rebuilding the managed codec and test assemblies.

- Four profiles: column, horizontal bar, line and secondary-axis categorical combo.
- Authored custom arrays and format grammar tokens survive native projection with the private PPJ snapshot removed. No-op preserves the exact ChartPart bytes.
- Value/format edits, both→plus→minus, custom→scalar→custom, deletion and recreation each verify native numLit values/formats, legal node order, Office 2021 series schema validation and fresh PPJ projection.
- Each source-bound edit changes only the selected errorBars owner; all other ChartML and ZIP-entry bytes remain unchanged.
- Invalid side combinations, scalar/custom contradictions, wrong lengths, negative/null/unsafe numbers, empty/control/oversized format strings, invalid format tokens and a forbidden formula property are rejected in both authored and imported compilation.
- Missing points, repeated indexes, extension nodes and negative native cache values withdraw analytics capability and preserve the original ChartPart on no-op.
- The existing source-bound guard regression now uses custom numRef/numCache formula owners: these remain unprojected and cannot be overwritten by a scalar insertion. Shared formula-presence and scalar/trendline lifecycle regressions also pass.

## Documentation and preview

Passed:

```sh
node skills/presentations/skills/presentation-skill-maintainer/scripts/maintain-presentation-skill.mjs sync
npm run docs:presentation-capabilities
node skills/presentations/skills/presentation-skill-maintainer/scripts/maintain-presentation-skill.mjs check
node test/ppj-preview-output-evidence.mjs
node test/reference-skill-sync.mjs
openspec validate ppj-chart-custom-error-data --strict
git diff --check
```

The existing preview diagnostic regression now includes custom plus/minus arrays for all four profiles and verifies `partial` / `chart-error-bars-not-rendered`. This checks truthful diagnostic behavior, not error-bar geometry.

`node test/skill-portability.mjs` was also run and still fails at the existing `installed REPL guidance: powerpoint-live-control` assertion. It expects REPL wording while the Live Skill documents typed `officekit live` commands; neither file is changed here.

## Remaining work

Formula-backed plus/minus data and cache/workbook synchronization remain an explicit F-07 gap. The projector does not flatten formula caches into literal PPJ fields. Unknown topology remains source-owned. Local SVG does not draw error bars. No full-suite, visual/host-PowerPoint or NativeAOT release acceptance is claimed.
