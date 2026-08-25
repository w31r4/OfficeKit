## ADDED Requirements

### Requirement: C route calibrates representative pages before full composition
For a deck longer than four pages, the C authoring route SHALL compose and render
an opening page, an evidence/data page, and the densest or highest-risk page
before completing the remaining pages.

#### Scenario: Long self-directed deck
- **WHEN** the authoring plan contains more than four pages
- **THEN** the Agent reviews the three representative pages together and may
  revise the design grammar before building the rest of the deck

#### Scenario: Short deck
- **WHEN** the authoring plan contains four pages or fewer
- **THEN** the full deck acts as one calibration spread and no separate partial
  deliverable is created

### Requirement: Calibration is internal and resumable
Calibration SHALL use the existing task, plan revision, Compose, render, and
review facilities without asking the user to select internal layouts or
creating a second design state.

#### Scenario: Grammar changes after calibration
- **WHEN** rendered evidence shows that the chosen grammar does not work across
  representative page types
- **THEN** the Agent writes a new authoring-plan revision and continues from that
  revision

#### Scenario: Fresh context resumes
- **WHEN** a new Agent context resumes the task after calibration
- **THEN** it can reopen the current plan and reviewed artifact state without
  restoring the previous JavaScript heap

### Requirement: Review evidence is independent of authoring rationale
The review stage SHALL inspect the plan, candidate, and rendered evidence and
MUST NOT treat the generator's written rationale as proof of successful visual
execution.

#### Scenario: Intent says chart but page shows generic text
- **WHEN** a page plan declares a chart as the dominant carrier but the rendered
  page does not show the intended data relationship
- **THEN** review reports the carrier mismatch even if the authoring notes claim
  the page is complete

### Requirement: Generalization check remains bounded and unseen
The change SHALL use three previously unused real briefs, run once each, and
SHALL NOT add them to normal, slow, package, or release gates.

#### Scenario: Dogfood exposes a repeatable failure
- **WHEN** an unseen brief reproduces a concrete product defect
- **THEN** the implementation fixes the defect and adds only the smallest
  regression assertion needed to prevent recurrence

#### Scenario: Dogfood produces a subjective preference
- **WHEN** reviewers merely prefer one valid visual direction over another
- **THEN** the result is recorded as judgment rather than converted into a hard
  design rule
