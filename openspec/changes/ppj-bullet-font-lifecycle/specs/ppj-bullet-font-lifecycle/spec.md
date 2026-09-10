## Purpose

Represent the direct font choice for a list marker and edit it without losing the source marker, surrounding text or unrelated document content.

## ADDED Requirements

### Requirement: Explicit font choice
PPJ SHALL accept either `text.paragraphs[].style.bullet.fontFamily` or `fontFollowText: true` for character, numbered and picture bullets. A family SHALL contain 1–255 XML-compatible Unicode scalars and non-whitespace content. Neither property means no direct font declaration. False, null, conflicting choices, malformed Unicode and invalid XML characters SHALL reject.

#### Scenario: Native round trip
- **WHEN** a valid family, follow-text choice or absent choice is authored and projected without embedded PPJ
- **THEN** PPJ and native marker font declarations retain the selected choice and distinguish absence from explicit follow-text

### Requirement: Independent source font lifecycle
Existing ordinary text/shape markers SHALL support font addition, assignment, switching, removal and restoration. Each changed property SHALL require its exact field authority, including both removed and added properties during a switch. The edit SHALL preserve marker content, other bullet styling, paragraph properties, runs, neighbors and non-target XML/ZIP. No-op SHALL retain source bytes.

#### Scenario: Set switch remove restore
- **WHEN** an authorized font edit is compiled from original source bytes and a fresh projection
- **THEN** only the target paragraph font declaration changes and native re-projection recovers the choice or absence

#### Scenario: Insufficient authority
- **WHEN** any changed font property lacks exact authority
- **THEN** compilation rejects without a candidate

### Requirement: Unknown font preservation
Unknown, malformed or duplicate native font choices SHALL remain source-owned. Unrelated edits SHALL preserve them and requested font replacement SHALL reject.

#### Scenario: Ambiguous imported font
- **WHEN** a source marker has duplicate font choices or unmodeled font attributes
- **THEN** its PPJ font choice is absent, no-op retains source bytes, unrelated edits preserve its XML and font replacement rejects

### Requirement: Discoverable font and preview boundary
Help, schema, registry, generated references and text guidance SHALL describe the same font choice and lifecycle. Preview SHALL retain existing direct character-font painting and explicitly diagnose unresolved follow-text/inherited font layout.

#### Scenario: Follow-text preview
- **WHEN** a character marker requests its font from text
- **THEN** preview reports the unresolved font/layout limitation instead of silently claiming full support
