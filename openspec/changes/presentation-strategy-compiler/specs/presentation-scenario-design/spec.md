## ADDED Requirements

### Requirement: Seven clean-room presentation scenarios
The Presentations Skill SHALL provide OfficeKit-authored guidance for analysis
and decision, business proposal, management report, academic research,
education and training, technical engineering, and brand creative work.

#### Scenario: Agent selects scenario guidance
- **WHEN** a new presentation request is routed
- **THEN** the Agent selects exactly one primary scenario, at most one secondary scenario, and reads the corresponding OfficeKit guide before composing pages

#### Scenario: Package remains clean-room
- **WHEN** the presentation Skill is packaged
- **THEN** it contains no copied Kimi reference text, Kimi filesystem path, or Kimi product identifier

### Requirement: Scenario and mechanism remain orthogonal
OfficeKit SHALL treat scenarios as communication contexts and existing design
mechanisms as generation approaches, allowing zero to two mechanisms under one
primary scenario.

#### Scenario: Technical management presentation
- **WHEN** a management report needs a technical system explanation
- **THEN** its primary scenario remains management-report while its mechanism packs can include enterprise-data-review and technical-architecture

### Requirement: Design authority precedence
User templates, brand systems, and explicit style references MUST override
scenario defaults; scenario guidance SHALL fill only decisions that the
authoritative source leaves unresolved.

#### Scenario: Template-conditioned deck
- **WHEN** a user supplies a presentation template and requests a management report
- **THEN** the template remains the design authority and the management-report guide affects narrative and unresolved composition choices without replacing template facts

### Requirement: Chosen direction precedes composition
For self-directed work, the Agent SHALL consider two or three internally
generated directions, persist one selected direction with rationale, and avoid
asking the user unless design authorities conflict materially.

#### Scenario: One-sentence creation
- **WHEN** a request supplies audience, purpose, and no conflicting design authority
- **THEN** the Agent chooses and records a direction before writing the design grammar and produces a complete working draft without a design questionnaire

### Requirement: Deck-specific design grammar
Every new plan SHALL define palette roles, surface hierarchy, typography,
geometry and line policy, density rhythm, visual carriers, media and chart
treatment, motif limits, and explicit anti-patterns.

#### Scenario: Avoid universal card styling
- **WHEN** an Agent authors several pages in one deck
- **THEN** it follows the selected grammar rather than applying one rounded rectangle, outline, and shadow treatment to every composition
