## ADDED Requirements

### Requirement: Presentation primitives shall be grouped by semantic job
OfficeKit MUST publish one concise primitive index grouping creation, visual
layers, text and typography, data/structure, motion, imported continuation, and
review/delivery operations by the Agent's intent rather than by source file.

#### Scenario: New Agent chooses a primitive
- **WHEN** an Agent reads the Presentation route for a page-level visual or
  imported-editing task
- **THEN** it can locate the relevant primitive family and its authoritative
  Help/API reference without loading unrelated format details

### Requirement: The primitive index shall state capability boundaries
Each indexed family MUST identify whether it is source-free, imported
source-bound, opaque-preserved, or fail-closed, and MUST link to the detailed
boundary when the operation is not universally editable.

#### Scenario: Agent targets an opaque imported object
- **WHEN** the index marks the object as opaque-preserved or fail-closed
- **THEN** the Agent is directed to inspect/capability evidence or refuse the
  mutation instead of inventing a raw OOXML path

### Requirement: The index shall not duplicate signatures
The primitive index MUST not become a second API reference; signatures,
parameter enumerations, and generated schemas remain owned by Help and API
documentation.

#### Scenario: Help changes an option
- **WHEN** an API option changes in the runtime
- **THEN** the semantic index only needs a link/boundary update if the Agent's
  choice or safety contract changes, rather than duplicating the full option
  table
