## Purpose

Provide independently editable paragraph default reflection values while preserving native optional state and surrounding source content.

## ADDED Requirements

### Requirement: Complete direct paragraph default reflection lifecycle

Ordinary text/shape paragraphs SHALL support defaultText.reflection assignment, deletion, single-effect defaultText/style wrapper removal and restoration under exact setTextParagraphStyle authority. All 14 reflection fields SHALL be optional: blur, distance, angle, startOpacity, endOpacity, startPosition, endPosition, fadeAngle, scaleX, scaleY, skewX, skewY, alignment and rotateWithShape. An empty object SHALL retain the effect with native defaults. Omission SHALL remain distinct from explicit zero, one and false.

Blur SHALL accept finite 0..1000pt and distance 0..100000pt, rounded to EMU. Angle and fadeAngle SHALL accept -360..360 degrees and wrap after native rounding. Opacity and position SHALL accept 0..1 at native 1/100000 precision; opacity token references SHALL require kind opacity. Scale ratios SHALL accept -21474.83648..21474.83647 at 1/100000 precision. Skew SHALL remain strictly within -90..90 degrees after native rounding. Alignment SHALL accept tl/t/tr/l/ctr/r/bl/b/br; rotateWithShape SHALL be boolean.

#### Scenario: Assign delete and restore optional values
- **WHEN** a valid reflection, any individual field, or a single-effect wrapper is assigned, removed and restored
- **THEN** native attributes and fresh PPJ preserve presence and normalized values while original direct runs, other paragraphs and non-target XML/ZIP remain unchanged

#### Scenario: Preserve mixed source effects
- **WHEN** reflection changes beside glow, inner shadow, outer shadow, soft edge or unknown sibling content
- **THEN** sibling values, source XML and effect-list attributes remain unchanged, and deleting reflection removes only its node and an empty attribute-free list

#### Scenario: Preserve unsupported source graphs
- **WHEN** the source contains duplicate reflection/list nodes, an effect DAG, or unknown/invalid reflection attributes or descendants
- **THEN** no-op preserves source bytes, unrelated default scalar edits retain that content, and reflection replacement rejects without output

#### Scenario: Reject invalid or unauthorized requests
- **WHEN** values exceed bounds, a rounded skew reaches its excluded endpoint, tokens are missing or have the wrong kind, authority is absent, or unrelated unsupported defaults are changed
- **THEN** compilation rejects without output and existing ordinary run/shape/image syntax constraints remain enforced
