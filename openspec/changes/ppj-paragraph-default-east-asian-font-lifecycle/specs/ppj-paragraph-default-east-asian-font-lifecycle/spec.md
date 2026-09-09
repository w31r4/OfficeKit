## Purpose

Provide direct East Asian paragraph default font editing and deletion through independently authorized PPJ fields while preserving other fonts and source state.

## ADDED Requirements

### Requirement: Direct East Asian font lifecycle

Ordinary text/shape paragraphs SHALL support independently authorized defaultText.fontFamilyEastAsia assignment, removal, font-only wrapper removal and restoration. Names SHALL be nonblank and at most 255 characters.

#### Scenario: Delete with a Latin font retained
- **WHEN** a source-bound request removes the direct East Asian font while retaining a Latin font
- **THEN** a:ea and freshly projected fontFamilyEastAsia are absent; the authoring fallback does not regenerate them, and other script fonts, effects, runs, paragraphs and ZIP entries remain unchanged

#### Scenario: Restore a font-only paragraph default
- **WHEN** an East-Asian-only defaultText/style wrapper is removed and an explicit East Asian font is later assigned
- **THEN** the direct native font and fresh PPJ projection follow the request without inventing Latin or complex-script fonts

#### Scenario: Reject unsupported replacement
- **WHEN** the name is invalid, field authority is missing, another unsupported field changes, or the source East Asian font contains extra metadata/child content
- **THEN** compilation rejects the request without candidate output and preserves source bytes
