## ADDED Requirements

### Requirement: Image commands are lazy and stable
OfficeKit SHALL expose `image search`, `image add`, `image list`, and `image audit` commands with stable JSON results, while root import, initialization, template search, and unrelated commands SHALL NOT load image providers, icon data, or perform network requests.

#### Scenario: Ordinary initialization remains lazy
- **WHEN** a user imports OfficeKit or runs `officekit init`
- **THEN** no image provider, Lucide collection, or remote request is initialized

#### Scenario: JSON command contract
- **WHEN** a valid image subcommand is invoked with `--json`
- **THEN** it emits one JSON result with stable status, task, command-specific data, and machine-readable errors

### Requirement: Search returns candidates without choosing
`image search` SHALL validate task, query, kind, purpose, orientation, and result limit; search only the explicitly selected v1 providers; record the search inside the task; and return candidates with task-bound opaque refs and `selectionMade: false`.

#### Scenario: Candidate discovery
- **WHEN** a photo query returns allowed Openverse and Wikimedia candidates
- **THEN** the CLI returns at most the requested count with source, preview, dimensions, license, attribution facts, and opaque candidate refs but no acquisition URL

#### Scenario: No candidate is valid
- **WHEN** every candidate has a blocked license or fails the requested orientation
- **THEN** search succeeds with an empty candidate list, rejection details, and `selectionMade: false`

#### Scenario: Provider error
- **WHEN** a requested provider fails
- **THEN** its failure is reported explicitly and OfficeKit does not silently substitute an undeclared provider

### Requirement: Acquisition accepts only explicit sources
`image add` SHALL acquire exactly one stored candidate, local file, or explicit HTTPS URL declaration and return the immutable task asset descriptor.

#### Scenario: Stored candidate acquisition
- **WHEN** a candidate ref belongs to the named task and remains valid
- **THEN** OfficeKit downloads the stored candidate URL under the secure acquisition policy and returns its absolute local path, MIME, dimensions, SHA-256, source, rights, and credit line

#### Scenario: Cross-task or unknown candidate
- **WHEN** a candidate ref is unknown or belongs to another task
- **THEN** acquisition fails without revealing or downloading its URL

#### Scenario: Explicit local file
- **WHEN** a user adds a supported local image with an allowed rights declaration
- **THEN** OfficeKit copies the bytes into task storage without modifying the source file

### Requirement: Audit reports actual PPTX media use
`image audit` SHALL hash media parts in the supplied PPTX and compare them with the named task's image receipts.

#### Scenario: Mixed registered and unregistered media
- **WHEN** a PPTX contains one registered task image and one unrelated embedded image
- **THEN** audit reports the registered asset as used, the unrelated media as unregistered, all unused registered assets, and outstanding attribution obligations

#### Scenario: Sources sidecar
- **WHEN** `--sources-output` names a distinct output path
- **THEN** OfficeKit atomically writes a deterministic JSON sidecar and never overwrites the PPTX input
