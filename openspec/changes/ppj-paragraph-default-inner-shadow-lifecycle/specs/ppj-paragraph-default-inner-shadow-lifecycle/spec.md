## Purpose

Provide independently editable paragraph default inner shadows with optional native geometry and opacity while preserving surrounding source content.

## ADDED Requirements

### Requirement: Paragraph default inner-shadow lifecycle

Ordinary text/shape paragraphs SHALL support defaultText.innerShadow assignment, deletion, inner-shadow-only defaultText/style wrapper removal and restoration under exact setTextParagraphStyle field authority. Color SHALL be required; blur, distance, angle and opacity SHALL preserve omission separately from explicit zero. Blur SHALL accept finite 0..1000pt, distance 0..100000pt and angle -360..360 degrees; lengths SHALL round to EMU, and angles SHALL normalize after native 1/60000-degree rounding. RGB/RGBA, theme/color tokens and opacity tokens SHALL retain their meanings, with declared grammar colors taking precedence in source edits and tint/shade resolving to RGB.

#### Scenario: Assign delete and restore optional state
- **WHEN** a valid inner shadow or any optional geometry/opacity field is assigned, removed and restored
- **THEN** native state and fresh PPJ preserve presence and normalized values while direct runs, other paragraphs and non-target XML/ZIP remain unchanged

#### Scenario: Preserve mixed effects
- **WHEN** a direct inner shadow is changed beside glow, outer shadow, reflection, soft edge or unknown sibling content
- **THEN** sibling values and source XML/attributes remain unchanged, and deletion removes only the target inner shadow and an empty attribute-free effect-list wrapper

#### Scenario: Retain unsupported source content
- **WHEN** the source has duplicate effect lists/inner shadows, an effect DAG or unmodeled inner-shadow geometry/color/descendants
- **THEN** no-op preserves source bytes, unrelated default scalar edits preserve the source content, and attempted inner-shadow replacement rejects without output

#### Scenario: Reject invalid or unauthorized input
- **WHEN** input has invalid values, missing/wrong-kind tokens, missing exact field authority or another unsupported default-field edit
- **THEN** compilation rejects without output; other text-style constraints and ordinary run/shape/image required-geometry syntax remain enforced
