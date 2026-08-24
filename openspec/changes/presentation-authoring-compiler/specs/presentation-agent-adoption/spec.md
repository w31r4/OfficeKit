## ADDED Requirements

### Requirement: Presentation Help adoption metadata
Each public Presentation Help entry SHALL have an `adoptionTier` of `golden`,
`advanced`, or `compatibility`. Golden entries SHALL additionally provide
bounded `useWhen`, `avoidWhen`, `requires`, `review`, and one or more packaged
recipe paths.

#### Scenario: Inspect a golden API
- **WHEN** Help returns `presentation.reuseSourceComponent`
- **THEN** the record explains when to use it, when not to use it, required inspected capability, review obligations, and a real packaged recipe

### Requirement: Intent search covers adoption fields
Help search SHALL index names, summaries, use/avoid guidance, capability
requirements, review guidance, and recipe labels without changing exact-name
lookup identity.

#### Scenario: Search by user intent
- **WHEN** an Agent searches for `reuse one component from an imported template`
- **THEN** Help ranks the bounded component resolve/reuse workflow ahead of unrelated low-level APIs

### Requirement: Golden API recipe coverage
The golden Presentation surface SHALL include Compose, AutoLayout, layout
validation, design profiling, template planning, inspect, component
resolve/edit/reuse, source-slide reuse, bounded SVG edits, and layout
placeholder discovery/use. Every golden API SHALL have a runnable example and
at least one black-box workflow assertion.

#### Scenario: Run the adoption audit
- **WHEN** the repository validates Presentation Help and Skill resources
- **THEN** no golden API has a missing recipe, missing example, missing file, or absent black-box assertion

### Requirement: Help remains the documentation source
Generated API documentation and packaged quick-start material SHALL derive
adoption facts from Help metadata. The change SHALL NOT introduce a separate
JSDoc example catalog that can disagree with Help.

#### Scenario: Regenerate API docs
- **WHEN** Help adoption metadata changes
- **THEN** generated documentation changes deterministically and repository checks reject stale output
