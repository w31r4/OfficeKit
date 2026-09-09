## Purpose

Preserve direct paragraph default italic presence during authorized source-bound PPJ editing without changing independent defaults or direct run formatting.

## ADDED Requirements

### Requirement: Default italic presence lifecycle
Ordinary text/shape paragraphs SHALL support true, false, deletion and restoration of text.paragraphs[].style.defaultText.italic.

#### Scenario: Edit and restore italic
- **WHEN** an authorized source-bound request changes or removes default italic
- **THEN** the direct native italic presence and fresh projection match the request while bold, other defaults, direct runs and unrelated XML/ZIP state remain unchanged.

#### Scenario: Remove an italic-only wrapper
- **WHEN** a defaultText or paragraph style wrapper containing only italic is removed
- **THEN** direct italic is removed and can be restored under the existing topology.

### Requirement: Independent authority
The codec SHALL require the exact italic field authority and reject unsupported style changes.

#### Scenario: Missing authority
- **WHEN** italic changes without its field authority or the request changes an unsupported default field
- **THEN** compilation is rejected without an output file.
