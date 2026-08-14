## ADDED Requirements

### Requirement: Applied Edit Plans are stored with reviewed task revisions
OfficeKit tasks SHALL store each applied plan and mutation footprint under the task's private `operations` directory and link it to the source and output revision hashes of the reviewed commit.

#### Scenario: Reviewed edit commit
- **WHEN** a FileBlob carrying a validated Edit Plan is committed with a passing review
- **THEN** OfficeKit writes one immutable operation record and links its hash and path from the task commit

#### Scenario: Failed or stale review
- **WHEN** the candidate review fails or its delivery hash is stale
- **THEN** no operation record advances task HEAD

### Requirement: Resume reconstructs from bytes rather than heap state
OfficeKit SHALL resume from the latest reviewed artifact revision, reimport its bytes, and rebuild node indices; it SHALL NOT claim to restore JavaScript objects or functions.

#### Scenario: Fresh Agent resumes a task
- **WHEN** a new process opens an existing task
- **THEN** it receives verified reviewed artifact bytes, prior operation evidence, and the next action needed to rebuild its semantic state

### Requirement: Operation records are bounded and tamper-evident
Operation records SHALL be schema-validated, size-bounded, written atomically with private permissions, content-addressed, and confined to the owning task.

#### Scenario: Corrupt operation record
- **WHEN** a stored operation record's bytes do not match its recorded SHA-256 or its path escapes the task root
- **THEN** task opening or commit resolution fails closed
