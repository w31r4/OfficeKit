# PPJ Preset Geometry Adjustments

## Why

PPJ already declares `shape.geometry.adjustments`, but authored build rejects
every non-empty value and imported PPTX projection drops recognized native
adjustments. This leaves common editable shapes such as rounded rectangles,
arrows, chevrons, stars, and callouts less expressive than the language claims.

## What Changes

- Define canonical ordered adjustment profiles for the PPJ preset geometries
  whose DrawingML guides can be represented as bounded literal values.
- Compile source-free PPJ adjustment arrays into native `a:avLst` guides.
- Project canonical imported adjustment guides back into PPJ.
- Issue a narrow source-bound geometry capability and lower only proven
  adjustment changes while preserving all unrelated package content.
- Keep angle-heavy, formula-valued, unknown, incomplete, or noncanonical
  adjustment graphs fail closed rather than exposing raw OOXML formulas.
- Replace the stale compiler-boundary entry with generated Agent guidance and
  one integrated round-trip/edit example.

## Capabilities

### New Capabilities

- `ppj-preset-geometry-adjustments`: Authored compilation, imported projection,
  source-bound editing, and Agent discoverability for canonical preset-shape
  adjustment arrays.

### Modified Capabilities

None.

## Impact

The PPJ schema, capability registry, generated `ppj.md`, protobuf wire-v2
messages, NativeAOT C# projector/compiler/codec, and one existing integrated
PPJ test are affected. The change is additive to wire v2 and does not expose
guide names, formulas, XPath, relationship IDs, or raw OOXML.
