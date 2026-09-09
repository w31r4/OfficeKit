## Purpose

Provide independently authorized direct paragraph language editing and deletion in source-bound PPJ while preserving run-level language and other source state.

## ADDED Requirements

### Requirement: Direct default language lifecycle

Ordinary text/shape paragraphs SHALL support defaultText.language assignment, deletion, language-only wrapper removal and restoration under exact setTextParagraphStyle field authority. The existing bounded tag grammar (2..63 characters, initial 2..8 letters, subsequent hyphen-separated 1..8 alphanumeric characters) SHALL apply without case normalization.

#### Scenario: Remove and restore a language
- **WHEN** a direct language or its language-only defaultText/style wrapper is removed and a valid language is later assigned
- **THEN** native lang and fresh PPJ presence follow the request while retaining altLang, fonts/effects, run languages, other paragraphs and non-target ZIP content

#### Scenario: Reject invalid or unauthorized values
- **WHEN** a language violates the bounded grammar, lacks field authority, or accompanies unsupported default-style edits
- **THEN** compilation rejects without output

#### Scenario: Preserve unmodeled source language
- **WHEN** source lang is outside the modeled grammar
- **THEN** no-op and unrelated scalar edits preserve it, and replacement through defaultText.language rejects without candidate output
