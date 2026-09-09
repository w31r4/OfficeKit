## Purpose

Allow agents to edit direct paragraph default bold in source-bound PPJ text and shapes while preserving explicit false, absence and unrelated formatting.

## ADDED Requirements

### Requirement: Direct paragraph default bold presence
Ordinary text and shape owners SHALL expose authorized editing of text.paragraphs[].style.defaultText.bold for fixed paragraph/run topology.

#### Scenario: Set remove and restore
- **WHEN** the field changes between true, false and absence
- **THEN** only its direct paragraph default bold state changes, fresh projection retains the distinction, and other defaults, direct runs and unrelated source parts remain unchanged.

#### Scenario: Remove a bold-only wrapper
- **WHEN** defaultText contains only bold and its field or wrapper is removed
- **THEN** the direct default bold is removed and can be restored without changing the text topology.

### Requirement: Exact capability and style boundary
The codec SHALL require the specific defaultText.bold paragraph-style capability and reject unrelated style or topology changes.

#### Scenario: Missing authority or unrelated edit
- **WHEN** the required capability field is absent, or a request also changes an unsupported defaultText field
- **THEN** compilation is rejected with no output file.
