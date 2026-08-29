## ADDED Requirements

### Requirement: The update Skill shall expose a deterministic impact map
The `skill-update` Skill MUST ship a versioned impact manifest that maps each
supported primitive family to its runtime owners, Help/API surface, semantic
reference, task route, examples, focused tests, and release evidence.

#### Scenario: Maintainer asks what a primitive change affects
- **WHEN** `skill-update impact` receives a changed runtime or protocol path
- **THEN** it prints the matching primitive families and the exact consumer
  paths that require review, without modifying any file

### Requirement: The checker shall fail on stale maintenance links
The checker MUST verify that manifest paths exist, referenced public Help names
are present, and direct Skill/reference links resolve within the repository.

#### Scenario: A primitive reference points to a removed API
- **WHEN** `skill-update check` reads a manifest entry whose Help name is absent
- **THEN** it exits non-zero and reports the family, missing name, and owning
  path

### Requirement: The update workflow shall remain lazy and host-neutral
The Skill and checker MUST use only repository text and git metadata. They MUST
NOT load Office WASM/NativeAOT, providers, browser tools, model context, or
network resources.

#### Scenario: Checker runs in a clean install
- **WHEN** `skill-update check` runs before any Office artifact import
- **THEN** it completes without initializing a codec, renderer, provider, or
  live host

### Requirement: The checker shall be advisory rather than an editor
The checker MUST report required follow-up work but MUST NOT rewrite Skill,
Help, API, source, or test files.

#### Scenario: Maintainer runs impact on a dirty worktree
- **WHEN** matching changes are present
- **THEN** the command reports the update checklist and leaves the worktree
  bytes unchanged
