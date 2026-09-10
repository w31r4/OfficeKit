## Purpose

Complete PPJ editing of direct paragraph left indent while retaining explicit zero, absence and surrounding imported content.

## ADDED Requirements

### Requirement: Complete paragraph indent lifecycle

Ordinary text/shape paragraphs SHALL support `text.paragraphs[].style.indent` as a finite number from 0 through 4032 points. Values SHALL round to the nearest EMU with ties to even. The existing paragraph-style priority SHALL select the value. Assignment, removal/restoration and removal of a style containing only indent SHALL round-trip through native PPTX and fresh PPJ. Explicit zero SHALL remain distinct from absence. Invalid types and out-of-range values SHALL reject without output. Every edit SHALL require the exact indent field authority.

#### Scenario: Assign, remove and restore
- **WHEN** an ordinary paragraph assigns zero, fractional or maximum indent, then removes and restores it
- **THEN** the native left margin and fresh PPJ SHALL retain the rounded value or absence
- **AND** missing indent authority SHALL reject the edit without output

### Requirement: Preserve source paragraph state

Indent edits SHALL change only the target paragraph's direct left margin. Hanging indent, spacing, unchanged native numeric spelling, other paragraph attributes/content, runs, neighbors and non-target XML/ZIP SHALL remain preserved. Invalid, negative or out-of-range native left-margin values SHALL remain source-owned: no-op preserves bytes, unrelated scalar edits preserve their XML, and replacement rejects without output.

#### Scenario: Preserve surrounding or unmodeled state
- **WHEN** an indent or unrelated scalar edit encounters other paragraph state or an unmodeled native left margin
- **THEN** only requested modeled state SHALL change and replacement of unmodeled left margin SHALL reject without output

### Requirement: Discoverability and preview evidence

Schema authority, Help, registry, generated references and text guidance SHALL agree that indent is a nonnegative paragraph left margin independent of hanging indent. Preview SHALL report its existing partial layout support, including explicit zero.

#### Scenario: Preview paragraph indent
- **WHEN** a paragraph declares indent, including zero
- **THEN** preview assessment SHALL expose a support-limit diagnostic without claiming host layout fidelity
