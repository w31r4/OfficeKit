## ADDED Requirements

### Requirement: Delivery-aware authoring plan
The presentation authoring plan SHALL accept optional `brief.deliveryMode` with
values `live`, `reader`, or `hybrid`, and `design.motionPolicy` with values
`adaptive`, `none`, or `explicit`. Missing values SHALL remain readable and
resolve to `hybrid` and `adaptive` for the C authoring route.

#### Scenario: Persist an explicit delivery mode
- **WHEN** an Agent writes a plan with `deliveryMode: "live"` and
  `motionPolicy: "adaptive"`
- **THEN** the canonical plan hash and task descriptor retain both values.

#### Scenario: Read a legacy plan
- **WHEN** a valid v1 plan has neither field
- **THEN** it remains readable without hash rewriting and the route reports the
  default hybrid/adaptive policy.

### Requirement: Structured composition intent
New C-route plans SHALL support a structured `pages[].compositionIntent` with a
visual `carrier`, `density`, `dominant`, and bounded `supporting` roles. Existing
non-empty string composition intents SHALL remain valid for migration.

#### Scenario: Select a visual carrier before motion
- **WHEN** a page declares a chart carrier and balanced density
- **THEN** the authoring route can select chart-specific composition and motion
  recipes without selecting an implicit Grid layout.

#### Scenario: Preserve legacy composition text
- **WHEN** a legacy page supplies a non-empty string composition intent
- **THEN** plan normalization succeeds and review may report a migration warning.

### Requirement: Bounded motion intent
A page MAY declare `motionIntent` containing one of six purposes, one matching
recipe, up to 32 ordered semantic units, and a transition of `none`, `fade`,
`push`, or `morph`. Unit IDs SHALL be unique and unit order SHALL be positive.

#### Scenario: Describe causal motion
- **WHEN** a page declares `causal-sequence` with ordered target roles and a
  `causal-reveal` recipe
- **THEN** the Agent can map the roles to typed animation calls and review can
  compare the emitted order to the plan.

#### Scenario: Reject an oversized motion intent
- **WHEN** a page declares more than 32 motion units or duplicate unit IDs
- **THEN** plan normalization fails without changing the active task plan.

### Requirement: C-route adaptive selection
The Presentations Skill SHALL select motion after narrative and composition. It
SHALL keep reader decks static, selectively animate hybrid decks, and allow
bounded live choreography while never adding sound or automatic advance.

#### Scenario: Reader deck stays static
- **WHEN** delivery mode is `reader` and the user did not explicitly request
  motion
- **THEN** the route emits no object animations and uses no automatic transition.

#### Scenario: Live deck uses a communication recipe
- **WHEN** delivery mode is `live` and a page has a chart or causal carrier
- **THEN** the route may select Data Rise or Causal Reveal and limits the page to
  a small ordered set of motion units.
