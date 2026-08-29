## ADDED Requirements

### Requirement: Presentation entry shall route progressively
The main Presentation Skill MUST contain only route selection, invariants, and
the common production spine; detailed primitives, scenarios, motion, imported
editing, and review rules MUST be loaded from one selected route or reference.

#### Scenario: Agent receives a new PPTX request
- **WHEN** the request is create, template-conditioned, imported edit,
  continuation, or review/delivery
- **THEN** the router selects exactly one task route before loading optional
  references

### Requirement: The route shall separate facts from guidance
The route MUST direct Agents to Help/API for observable runtime facts and to
design/scenario references for choices, without presenting examples as a
universal visual template.

#### Scenario: Agent needs an animation option
- **WHEN** the Agent is composing a live deck
- **THEN** it loads the motion reference and queries Help/API for the supported
  call, rather than copying an unrelated example's palette or helper

### Requirement: The route shall preserve source-bound safety
Progressive loading MUST NOT weaken input immutability, capability checks,
opaque preservation, second import, review, or no-overwrite delivery rules.

#### Scenario: Agent edits an imported slide
- **WHEN** the selected operation has no issued capability
- **THEN** the route sends the Agent to inspect/fail-closed guidance and does
  not suggest raw package mutation
