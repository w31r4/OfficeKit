## Purpose

Make local PPJ SVG/PNG preview output non-destructive and auditable, so users can distinguish files actually produced from missing or failed artifacts and associate them with the exact inputs used.

## ADDED Requirements

### Requirement: Exclusive preview destination

Preview publication SHALL require a new destination directory, reject existing files, directories and symlinks at that destination, and never overwrite the input PPJ, source package, assets or previous evidence. Every artifact write SHALL be exclusive. Page identifiers SHALL NOT escape the destination or cause output filename collisions, including on case-insensitive filesystems.

#### Scenario: Existing destination

- **WHEN** preview targets an existing directory, ordinary file or symlink
- **THEN** publication fails without changing the destination or any input bytes

#### Scenario: Concurrent publication

- **WHEN** two previews attempt to create the same destination
- **THEN** at most one acquires it and the other fails without overwriting the first run

#### Scenario: Unsafe or colliding page identifiers

- **WHEN** page identifiers contain path separators, reserved filename forms or case-insensitive collisions
- **THEN** output filenames are safely mapped or rejected before writing page artifacts, and each original page ID remains traceable

### Requirement: Explicit incomplete output lifecycle

After reserving an output directory, preview SHALL label it as incomplete before writing page artifacts. It SHALL record final output status only after all requested artifact attempts are resolved. Recoverable failures SHALL preserve successful artifacts and record the failed stage/page without presenting incomplete output as complete. If final evidence cannot be written, the run SHALL fail and SHALL NOT leave a successful completion record.

#### Scenario: Raster failure after one successful page

- **WHEN** PNG production fails after another page has been written
- **THEN** the successful artifacts remain, the failed page has no claimed PNG artifact, final evidence reports incomplete output, and the operation fails

#### Scenario: Evidence write failure or interrupted publication

- **WHEN** publication cannot finish the final manifest or terminates before it is published
- **THEN** readers do not observe a complete final result and any retained pending marker identifies the run as incomplete

### Requirement: Manifest describes actual artifacts

The final manifest SHALL carry a versioned schema, original page IDs, actual artifact filenames, SHA-256 hashes and byte lengths computed from the bytes published. Compatibility page-level SVG/PNG references SHALL be included only for files that were successfully written. A valid final manifest SHALL distinguish output production status from visual correctness and SHALL NOT claim human review or editing fidelity.

#### Scenario: Successful SVG and PNG publication

- **WHEN** all requested artifacts are written
- **THEN** every manifest artifact can be read and its hash and size match the manifest, output status is complete, and visual correctness remains separately unevaluated

#### Scenario: Artifact write failure

- **WHEN** an artifact cannot be written
- **THEN** it is absent from the actual-artifact list and compatibility file references, and the failed attempt is represented separately

### Requirement: Missing raster dependency is not successful PNG output

When output publication requests the existing default SVG plus PNG artifacts, an unavailable raster dependency SHALL retain obtainable SVGs, report the dependency as unavailable, omit nonexistent PNG references and fail the operation with incomplete output evidence. Preview SHALL NOT download a dependency or switch to another authoring/rendering engine automatically.

#### Scenario: Raster dependency unavailable

- **WHEN** PNG generation is requested but its dependency cannot be loaded
- **THEN** generated SVGs remain inspectable, the evidence lists no unproduced PNG, and the result cannot be interpreted as successful PNG publication

### Requirement: Input and asset identity matches rendered bytes

Evidence SHALL identify the input path and raw input hash, canonical program hash, compiled candidate hash, source-bound flag and source hash when present, and the actual asset IDs/hashes/byte sizes used. The assets embedded in the preview SHALL come from the same loaded byte snapshot supplied to compilation, not a later read of mutable paths. Evidence SHALL identify the preview implementation and loaded raster/runtime versions, recording unavailable identity explicitly rather than inventing it.

#### Scenario: Asset file changes after load

- **WHEN** an asset path is changed after the input snapshot was loaded
- **THEN** compilation and preview both use the loaded bytes and the evidence hash describes those bytes

#### Scenario: Source-bound input

- **WHEN** an imported PPJ with source bytes is previewed
- **THEN** source identity is recorded without modifying or flattening its original package, and the candidate identity refers to the compiled result rather than the original source by assumption

### Requirement: In-memory preview remains lazy

An invocation that does not request output publication SHALL NOT create output directories or load the PNG dependency solely for publication. Adding output evidence SHALL NOT initialize rendering dependencies through the root package import.

#### Scenario: In-memory SVG preview without raster dependency

- **WHEN** a caller requests only the existing in-memory preview result
- **THEN** it can obtain SVG results without publication files or a raster dependency, subject to normal input compilation checks
