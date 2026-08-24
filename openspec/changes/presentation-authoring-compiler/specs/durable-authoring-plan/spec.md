## ADDED Requirements

### Requirement: Bounded authoring-plan revisions
The task REPL SHALL expose one `ctx.plan([value], options)` surface for the
`office-kit/presentation-authoring-plan/v1` plain-JSON schema. It SHALL reject
unsupported schemas, non-JSON values, raw OOXML fields, values over 256 KiB,
more than 64 ordered pages, and source/template references that are not bound
to a task artifact ID and SHA-256.

#### Scenario: Store and read a plan
- **WHEN** an Agent writes a valid presentation authoring plan
- **THEN** OfficeKit writes one private immutable hash-addressed plan revision, atomically selects it, and returns the same defensive data from `ctx.plan()`

#### Scenario: Reject unsafe plan state
- **WHEN** a plan contains a function, cyclic value, unsupported schema, raw OOXML selector, excessive bytes, or excessive pages
- **THEN** the write fails without changing the active plan descriptor

### Requirement: Optimistic and idempotent plan updates
An update SHALL require the exact current plan SHA-256 after a plan exists.
Writing identical canonical content SHALL be idempotent and SHALL NOT create a
new revision or update timestamp.

#### Scenario: Reject a stale update
- **WHEN** an Agent writes a changed plan with a missing or stale `expectedSha256`
- **THEN** OfficeKit rejects it and preserves the current plan

#### Scenario: Repeat the same plan
- **WHEN** an Agent writes canonical content equal to the active plan with its exact expected hash
- **THEN** OfficeKit returns the existing descriptor without adding another revision

### Requirement: Reviewed plan and artifact stay bound
A Presentation review using an authoring plan SHALL record its SHA-256. A task
commit SHALL snapshot the active plan and SHALL reject a review made against a
different plan. Publishing SHALL fail while the active plan is newer than the
current reviewed commit.

#### Scenario: Commit a planned presentation
- **WHEN** the candidate hash, review delivery hash, review plan hash, and active task plan all agree
- **THEN** the commit snapshots that plan descriptor and becomes publishable

#### Scenario: Change intent after review
- **WHEN** an Agent updates the plan after the latest reviewed commit
- **THEN** the task reports a working plan and refuses publication until a new reviewed commit binds it

### Requirement: Task v1 read compatibility and lazy migration
Task readers SHALL accept schema-1 manifests as planless tasks. Listing and
detail commands SHALL NOT rewrite them. The first successful mutating task
operation SHALL atomically write schema 2 without altering existing inputs,
revisions, reviews, operations, publications, or commit IDs.

#### Scenario: List an old task
- **WHEN** `officekit tasks` reads a valid schema-1 task
- **THEN** it reports `plan: null` without modifying task bytes or loading any artifact runtime

#### Scenario: Mutate an old task
- **WHEN** a schema-1 task successfully stages input, writes a plan, commits, or publishes
- **THEN** it becomes schema 2 while all prior task facts remain unchanged

### Requirement: Compact task and resume presentation
Task detail and the protocol-3 ready envelope SHALL expose a bounded plan
descriptor containing schema, mode, page count, recipe, state, SHA-256, bytes,
and managed path. They SHALL NOT inline full plan content in task listings.

#### Scenario: Resume a planned deck
- **WHEN** a fresh process opens a task with an active authoring plan
- **THEN** the ready envelope identifies the plan and `ctx.plan()` revalidates and returns its full content without restoring JavaScript heap state
