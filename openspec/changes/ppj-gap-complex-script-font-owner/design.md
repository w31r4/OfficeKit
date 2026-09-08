# Design

## Contract

- `textStyle.fontFamilyComplexScript` and run-level text style values are
  bounded literal or string grammar-token typefaces.
- The native wire stores the direct typeface separately from Latin and East
  Asian faces.
- Authored PresentationML writes one `a:cs/@typeface` under the relevant
  `a:rPr` or `a:defRPr`.
- PPJ projection emits the field when the imported native child is a simple
  typeface-only `a:cs` element.
- A source-bound `fontFamilyComplexScript` native leaf may replace only that
  existing direct run child after the usual revision, semantic, and topology
  proofs.

## Boundaries

The slice does not infer a complex-script face from language, select a font
from the host, rewrite theme `a:fontScheme`, or flatten a complex `a:cs`
extension/effect graph. Unsupported or ambiguous native children remain
opaque and fail closed.
