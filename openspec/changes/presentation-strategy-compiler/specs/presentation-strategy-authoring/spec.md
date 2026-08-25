## ADDED Requirements

### Requirement: Communication-first authoring plans
OfficeKit SHALL accept additive communication-job, expected-outcome,
medium-fit, after-use, scenario, and chosen-direction fields in presentation
authoring-plan v1 without changing its schema identifier.

#### Scenario: New plan records presentation strategy
- **WHEN** an Agent writes a new presentation authoring plan
- **THEN** the plan records one primary communication job, an expected audience change, one primary scenario, and one chosen visual direction

#### Scenario: Weak medium fit remains actionable
- **WHEN** the Agent determines that PowerPoint is a weak fit for a user-requested deliverable
- **THEN** it records the limitation and continues the requested presentation workflow without forcing another user decision

### Requirement: Existing tasks remain readable
OfficeKit MUST continue to read authoring-plan v1 tasks that predate the
strategy fields and MUST report their missing strategy as a compatibility
warning rather than rejecting or migrating them.

#### Scenario: Resume legacy task
- **WHEN** a task with a valid legacy authoring-plan v1 is resumed
- **THEN** the plan remains readable and its strategy descriptor fields are absent or null without changing the stored revision

### Requirement: Strategy survives task continuation
Task and REPL summaries SHALL expose the active plan's communication job,
primary scenario, chosen direction, delivery mode, and pending strategy issues.

#### Scenario: Fresh context resumes a deck
- **WHEN** a new Agent resumes a strategy-bound presentation task
- **THEN** it can recover the selected communication job, scenario, direction, reviewed revision, and next action without restoring the previous JavaScript heap

### Requirement: Public presentation doctrine
OfficeKit SHALL publish an English canonical presentation doctrine and a
synchronized Chinese edition that define the medium, lifecycle, quality
layers, and OfficeKit responsibility boundary.

#### Scenario: Reader follows the product doctrine
- **WHEN** a user follows the README presentation-positioning link
- **THEN** the doctrine explains factual, communication, narrative, cognitive, visual, and native/run-time quality without promising automated judgment of facts or aesthetics
