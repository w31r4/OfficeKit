## Purpose

Allow direct paragraph default Latin font assignment and deletion through independently authorized PPJ fields while retaining source text and other style state.

## ADDED Requirements

### Requirement: Default font family has a complete direct lifecycle

Ordinary text/shape paragraphs SHALL support defaultText.fontFamily assignment, deletion, font-only wrapper removal and restoration under the exact setTextParagraphStyle field authority. Names SHALL be nonblank and at most 255 characters.

#### Scenario: Remove and restore a simple font
- **WHEN** the directly modeled family or its font-only defaultText/style wrapper is removed and a family is later assigned
- **THEN** native Latin font presence and fresh PPJ projection follow each request while retaining other script fonts, effects, direct runs, other paragraphs and unrelated ZIP entries

#### Scenario: Reject invalid or unauthorized edits
- **WHEN** a family is invalid, field authority is missing, or unrelated unsupported style changes accompany it
- **THEN** compilation rejects the request without candidate output

#### Scenario: Preserve unmodeled native font content
- **WHEN** a source Latin font has additional metadata or child content outside the modeled typeface field
- **THEN** replacement is rejected and source bytes remain unchanged rather than flattening the native font
