# PPJ authored diagrams

## Why

PPJ already declares `smartArt.mode: "authored"`, a bounded node graph, and
eight semantic layout families, but the authored compiler rejects every such
element. This is a real language/compiler gap: an Agent can describe a process,
hierarchy, cycle, relationship, matrix, pyramid, or picture sequence, yet PPJ
cannot lower that intent into editable PowerPoint objects.

## What changes

- Give authored diagram nodes explicit shape, text, and optional picture style
  state instead of relying on implicit compiler aesthetics.
- Lower authored diagrams deterministically to one native editable
  `PresentationGroup` containing ordinary shapes, connectors, text, and images.
- Support list, process, cycle, hierarchy, relationship, matrix, pyramid, and
  picture layouts with stable generated child IDs and bounded node counts.
- Preserve source-bound SmartArt as nativeRef-backed content; do not rebuild a
  third-party SmartArt graph into authored shapes.

## What does not change

- OfficeKit does not claim that the authored lowering is a native Office
  SmartArt diagram part. It is a semantic PPJ diagram compiled to editable
  DrawingML objects.
- PPJ does not expose SmartArt OOXML, layout URIs, relationship IDs, or arbitrary
  layout algorithms.
- The compiler does not choose colors, fonts, or decorative geometry. The PPJ
  program supplies named styles and picture assets.

