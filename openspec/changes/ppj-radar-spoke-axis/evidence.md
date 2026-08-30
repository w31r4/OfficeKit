# Evidence

## Implemented contract

- PPJ radar charts accept one semantic `spokeAxis` for visibility, numeric
  bounds and interval, label visibility/number format/direct typography,
  radial spoke style and concentric-ring style.
- NativeAOT lowers the object to the canonical standard-radar category/value
  axis pair and reconstructs it from a snapshot-free canonical ChartPart.
- Recognized imported radar charts issue bounded axis and text-style
  capabilities. A local edit patches the existing ChartPart, preserves its
  topology and reprojects the requested semantic state.
- Unsupported custom label positions, reverse/logarithmic radar axes, arrows,
  theme/effect line graphs, filled/3D variants, extensions and irregular axis
  topology remain source-owned and fail closed.

## Lean verification

The implementation extended the existing comprehensive PPJ contract rather
than adding a radar matrix or new fixture file.

- `dotnet test native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj -p:UseSharedCompilation=false --filter 'FullyQualifiedName~PpjV1CompilesCanonicalPresentationProgramDeterministically' --no-restore --logger 'console;verbosity=minimal'`: passed, 1 of 1 focused tests.
- The focused contract covers schema conflict rejection, authored native XML,
  snapshot-free semantic projection, source-bound max/label/ring edits,
  post-write semantic proof and second projection.
- Agent discoverability is checked through the capability registry, generated
  PPJ manual and chart guidance.

Not run: full `npm test`, package/release gates, visual rendering, Keynote or
Windows PowerPoint. Those remain PPJ 2.0 release-level evidence rather than a
claim of host-rendered radar parity.
