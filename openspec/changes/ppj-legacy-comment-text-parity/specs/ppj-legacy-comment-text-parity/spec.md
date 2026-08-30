## ADDED Requirements

### Requirement: Imported comment binding
Every projected bounded legacy comment SHALL carry a source-revision-bound
nativeRef, and editable profiles SHALL issue `replaceText` for `text`.

#### Scenario: Editable legacy comment
- **WHEN** one slide owns a supported legacy comments part and the shared author
  catalog is closed and relationship-free
- **THEN** its PPJ comment advertises text replacement capability

### Requirement: Fixed-topology comment text edit
A capable source-bound comment change SHALL alter only its text while retaining
all comment identity and topology state.

#### Scenario: Revise one review comment
- **WHEN** an Agent changes only one PPJ comment's `text`
- **THEN** build changes only the owning comments part and reimport recovers the
  requested text

### Requirement: Other comment mutations fail closed
Source-bound PPJ SHALL reject comment add/remove/reorder and changes to page,
target, parent, author, date, resolved state, position, IDs, or nativeRef.

#### Scenario: Move a comment
- **WHEN** an Agent changes comment position together with its text
- **THEN** build rejects the mutation before native export

### Requirement: Agent discoverability
PPJ review guidance SHALL distinguish authored comment features from the narrow
imported legacy text capability.

#### Scenario: Agent responds to imported review feedback
- **WHEN** an Agent reads the comment guidance
- **THEN** it can determine whether the exact comment is text-editable and knows
  that author, position, topology, and resolution remain fixed
