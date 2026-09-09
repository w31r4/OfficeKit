## Purpose

Complete the direct complex-script paragraph default font lifecycle in source-bound PPJ while preserving other script fonts and unrelated source content.

## ADDED Requirements

### Requirement: Independent complex-script default font presence

Ordinary text/shape paragraphs SHALL support defaultText.fontFamilyComplexScript assignment, deletion, font-only wrapper removal and restoration with exact setTextParagraphStyle field authority. Names SHALL be nonblank and at most 255 characters.

#### Scenario: Delete and restore a direct font
- **WHEN** the complex-script font or its font-only defaultText/style wrapper is removed and an explicit font is later restored
- **THEN** native a:cs presence and fresh PPJ projection follow the request while retaining Latin/East Asian fonts, other defaults/effects, runs, paragraphs and non-target ZIP entries

#### Scenario: Reject unsupported or unauthorized edits
- **WHEN** a name is invalid, authority is missing, another unsupported field changes, or the source a:cs contains extra metadata or child content
- **THEN** compilation rejects without output and source bytes remain unchanged
