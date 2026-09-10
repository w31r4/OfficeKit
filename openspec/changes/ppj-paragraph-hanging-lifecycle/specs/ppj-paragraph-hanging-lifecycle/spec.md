## Purpose

Complete PPJ editing of direct paragraph hanging indent while retaining signed values, explicit zero, absence and imported source content.

## ADDED Requirements

### Requirement: Complete hanging-indent lifecycle

Ordinary text/shape paragraphs SHALL support `text.paragraphs[].style.hanging` as a finite number from -4032 through 4032 points, independent of left indent. Positive values SHALL represent leftward first-line hanging, and negative values rightward first-line indentation; native indentation SHALL carry the opposite sign. Values SHALL round to the nearest EMU with ties to even. Existing paragraph-style priority SHALL select the value. Assignment, removal/restoration and removal of a style containing only hanging SHALL round-trip through native PPTX and fresh PPJ. Explicit zero SHALL remain distinct from absence. Invalid types and out-of-range values SHALL reject without output. Every edit SHALL require exact hanging field authority.

#### Scenario: Assign, remove and restore
- **WHEN** an ordinary paragraph assigns signed, zero, fractional or endpoint hanging values, then removes and restores the field
- **THEN** native indentation and fresh PPJ SHALL retain the rounded signed value or absence
- **AND** missing hanging authority SHALL reject without output

### Requirement: Preserve source paragraph state

Hanging edits SHALL change only the target paragraph's direct native indentation. Left indent, spacing, unchanged native numeric spelling, other paragraph attributes/content, runs, neighbors and non-target XML/ZIP SHALL remain preserved. Invalid or out-of-range native indentation SHALL remain source-owned: no-op preserves bytes, unrelated scalar edits preserve its XML, and replacement rejects without output.

#### Scenario: Preserve surrounding or unmodeled state
- **WHEN** a hanging or unrelated scalar edit encounters other paragraph state or unmodeled native indentation
- **THEN** only requested modeled state SHALL change and unmodeled replacement SHALL reject without output

### Requirement: Discoverability and preview evidence

Schema authority, Help, registry, generated references and text guidance SHALL agree on hanging's sign, range and independent lifecycle. Preview SHALL report existing partial layout support, including explicit zero.

#### Scenario: Preview hanging indent
- **WHEN** a paragraph declares hanging, including zero
- **THEN** preview assessment SHALL expose a support-limit diagnostic without claiming host layout fidelity
