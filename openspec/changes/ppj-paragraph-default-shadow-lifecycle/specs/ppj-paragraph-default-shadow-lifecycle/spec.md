## Purpose

Provide full direct paragraph default outer-shadow values with independent lifecycle editing and preservation of surrounding source content.

## ADDED Requirements

### Requirement: Paragraph default outer-shadow lifecycle

Ordinary text/shape paragraphs SHALL support defaultText.shadow assignment, deletion, shadow-only defaultText/style wrapper removal and restoration under exact setTextParagraphStyle authority. Color SHALL be required; opacity, blur, distance, angle, alignment, rotateWithShape, scaleX, scaleY, skewX and skewY SHALL be optional. Omission SHALL remain distinct from explicit zero, one and false.

Blur SHALL accept finite 0..1000pt and distance 0..100000pt, rounded to EMU. Angle SHALL accept -360..360 degrees and wrap after native rounding. Opacity SHALL accept 0..1 or a token of kind opacity at 1/100000 precision. Scale ratios SHALL accept -21474.83648..21474.83647 at 1/100000 precision. Skew SHALL remain strictly within -90..90 degrees after native rounding. Alignment SHALL accept tl/t/tr/l/ctr/r/bl/b/br and rotateWithShape SHALL be boolean. RGB/RGBA and color tokens SHALL retain their meanings; simple source theme colors SHALL retain scheme identity, declared source grammar colors SHALL take precedence and tint/shade SHALL resolve to RGB.

#### Scenario: Assign remove and restore shadow values
- **WHEN** a valid shadow or any optional value is assigned, removed and restored
- **THEN** native attributes and fresh PPJ preserve presence and normalized values, including a color-only shadow, while other paragraphs, direct runs and non-target XML/ZIP remain unchanged

#### Scenario: Preserve mixed effects and wrapper state
- **WHEN** shadow changes beside glow, inner shadow, reflection, soft edge or unknown sibling content, or a shadow-only wrapper is removed
- **THEN** only the target shadow node and an empty attribute-free list are removed or replaced, and sibling values/XML and effect-list attributes are retained

#### Scenario: Retain unsupported source content
- **WHEN** the source contains duplicate shadow/list nodes, an effect DAG, invalid/unknown shadow attributes or unmodeled color descendants
- **THEN** no-op preserves source bytes, unrelated default scalar edits retain the content, and shadow replacement rejects without output

#### Scenario: Reject invalid or unauthorized requests
- **WHEN** color is missing, values exceed bounds, rounded skew reaches an excluded endpoint, tokens are missing or wrongly typed, authority is absent or another unsupported default field is changed
- **THEN** compilation rejects without output while ordinary run/shape/image required-geometry syntax remains enforced
