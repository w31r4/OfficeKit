# Evidence

## Implemented contract

- Ordinary PPJ chart axes accept independent bounded start/end arrowheads.
- NativeAOT writes and reads canonical DrawingML `a:headEnd` and `a:tailEnd`
  without introducing overlay connectors.
- Canonical imported axes project semantic endpoint names and accept
  capability-issued add/replace/remove edits inside the existing ChartPart.
- Grid, series, marker, trend, error-bar, radar-spoke and generated vector-axis
  arrows remain unsupported; endpoint sizing/effects and irregular line graphs
  stay source-owned.

## Lean verification

The existing comprehensive PPJ contract was extended rather than adding an
arrow matrix or a new fixture.

- `dotnet build native/OfficeKit/src/OfficeKit.Codec/OfficeKit.Codec.csproj -p:UseSharedCompilation=false --no-restore --nologo`: passed with 0 warnings and 0 errors.
- `dotnet test native/OfficeKit/tests/OfficeKit.Codec.Tests/OfficeKit.Codec.Tests.csproj -p:UseSharedCompilation=false --filter 'FullyQualifiedName~PpjV1CompilesCanonicalPresentationProgramDeterministically' --no-restore --logger 'console;verbosity=minimal'`: passed, 1 of 1 focused tests.
- The focused contract proves hidden-line conflict rejection, authored
  open/triangle native XML, snapshot-free PPJ recovery, source-bound
  none/diamond editing, post-write semantics and second projection.

Not run: full `npm test`, package/release gates, visual rendering, Keynote or
Windows PowerPoint. Those remain PPJ 2.0 release-level evidence rather than
host-rendered arrowhead evidence.
