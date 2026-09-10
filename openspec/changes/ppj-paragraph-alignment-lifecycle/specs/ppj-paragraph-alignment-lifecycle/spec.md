## Purpose

Let PPJ users edit direct paragraph alignment while retaining source content and the difference between an explicit value and inherited alignment.

## ADDED Requirements

### Requirement: Direct paragraph alignment lifecycle

Ordinary source-bound text and shape owners SHALL support independent assignment, removal and restoration of `text.paragraphs[].style.alignment` under its exact field authority. The five values SHALL remain left, center, right, justify and distributed. Explicit left SHALL retain a native attribute; omission or removal of an alignment-only style wrapper SHALL remove it. Direct authored paragraph alignment SHALL override the owner's paragraph default.

#### Scenario: Edit and reproject a direct alignment
- **WHEN** a user changes, removes or restores the alignment of one paragraph
- **THEN** fresh projection SHALL retain the requested value or absence, preserving runs, neighboring paragraphs, other paragraph properties and non-target XML/ZIP content

#### Scenario: Invalid or unauthorized request
- **WHEN** a request uses an invalid value or lacks alignment field authority
- **THEN** compilation SHALL reject it without an output file

### Requirement: Preserve unmodeled native alignment

Unmodeled native alignment values SHALL remain source-owned, omitted from modeled PPJ, preserved by no-op and unrelated scalar edits, and protected against alignment replacement.

#### Scenario: Unknown source alignment
- **WHEN** an imported paragraph contains an unsupported or invalid native alignment token
- **THEN** unrelated default-bold assignment/removal SHALL preserve it, and an alignment replacement SHALL fail without output

### Requirement: Honest capability documentation

Help, schema, capability metadata and Agent references SHALL describe the direct field lifecycle and preserve explicit limits on preview and host text layout.

#### Scenario: Discover and preview the field
- **WHEN** a user reads generated PPJ guidance or previews a paragraph with alignment
- **THEN** the five field values and lifecycle SHALL be discoverable, and incomplete preview layout SHALL remain explicitly partial
