# Design

## Owner and boundary

- Owner: `p:sp/p:txBody/a:bodyPr/a:prstTxWarp/a:avLst/a:gd/@fmla`.
- PPJ semantic location: `textBoxStyle.textWarpAdjustments`.
- Native leaf: `textBodyWarpAdjustment`.
- Identity is the ordered `a:gd` index inside the one direct `a:avLst`; the
  guide name is carried in the semantic field but is not an arbitrary XML
  selector supplied by callers.
- A valid adjustment list has one direct `prst` attribute, one optional
  `avLst` child, unique non-empty `name` attributes, one `fmla` attribute per
  guide, and a canonical `val` followed by a signed 32-bit integer. No guide
  has children or extension attributes.
- The existing `textWarpPreset` projection is extended only for this same
  recognized topology. A source-bound adjustment edit changes one formula
  value token and leaves the preset, guide names/order, text, effects, and
  other package parts untouched.

## Evidence

The focused test authors a `textArchUp` text body with one `adj` value, checks
the literal `a:avLst/a:gd` XML and Open XML validity, removes embedded PPJ,
projects the preset, adjustment field, and indexed native leaf, then edits the
leaf to a new integer. It verifies that only `ppt/slides/slide1.xml` changes,
the other ZIP parts remain byte-identical, and a second projection recovers the
new value. Empty, formula-driven, extension-bearing, and duplicate-name lists
do not receive the bounded adjustment leaf.
