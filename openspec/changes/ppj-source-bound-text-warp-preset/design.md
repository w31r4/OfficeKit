# Design

## Owner and boundary

- Owner: `p:sp/p:txBody/a:bodyPr/a:prstTxWarp/@prst`.
- PPJ semantic location: `textBoxStyle.textWarpPreset`.
- Native leaf: `textBodyWarpPreset`.
- The value is the finite DrawingML `TextShapeValues` token.
- The additive `PresentationTextBodyProperties.text_warp_preset = 37` field
  keeps protocol 2 stable.
- A single preset child is modeled when it is bare or carries one recognized
  literal `a:avLst/a:gd` list. Extensions, duplicate preset children, unknown
  values, and non-literal adjustment formulas remain source-owned.
- A source-bound edit replaces the existing `prst` attribute value only; it
  does not create a missing child or rewrite adjustment formulas.

## Evidence

The focused test authors a text body with `textArchUp`, removes the embedded
PPJ, projects one native leaf, changes it to `textPlain`, and verifies that
only `ppt/slides/slide1.xml` changes, Open XML remains valid, all other ZIP
parts remain byte-identical, and the second projection recovers the new
preset. It also verifies that a preset with an adjustment child and a missing
preset do not receive the leaf.
