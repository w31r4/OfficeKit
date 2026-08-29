## ADDED Requirements

### Requirement: Typography guidance shall be role-based
OfficeKit MUST provide a compact typography reference that chooses fonts by
role, language, medium, scenario, and evidence of installed/renderable fonts.

#### Scenario: Agent creates a Chinese analytical deck
- **WHEN** the deck includes CJK body text and English data labels
- **THEN** the Agent selects an explicitly compatible CJK/Latin pairing and
  records the fallback or render limitation rather than relying on a guessed
  font name

### Requirement: Typography guidance shall preserve design freedom
The reference MUST describe selection criteria and anti-patterns, not impose a
single font family, palette, page silhouette, or template.

#### Scenario: User supplies a design system
- **WHEN** the user's font rules conflict with the default guidance
- **THEN** the user design system wins and the reference supplies only missing
  compatibility or fallback advice

### Requirement: Typography changes shall be visible to review
The Presentation route MUST direct Agents to inspect rendered title, body,
label, source, and CJK fallback behavior before delivery.

#### Scenario: Installed font is substituted
- **WHEN** rendering changes a title's metrics or causes overflow
- **THEN** the Agent records the substitution and adjusts content or font choice
  before calling the deck final
