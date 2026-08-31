## ADDED Requirements

### Requirement: Task images are immutable and resumable
OfficeKit SHALL store image bytes by SHA-256 beneath the selected task, write private provenance receipts and search evidence, and recover them through task open/list operations without changing the task manifest schema.

#### Scenario: Idempotent duplicate add
- **WHEN** the same supported image bytes are added twice with compatible provenance
- **THEN** OfficeKit reuses one content-addressed asset and returns the same SHA-256 and path

#### Scenario: New context resumes image evidence
- **WHEN** a later Agent opens the same task
- **THEN** image list and stored candidate refs remain discoverable without repeating the provider search

#### Scenario: Unsafe task path
- **WHEN** a task asset or evidence path would escape the task root or traverse a symbolic link
- **THEN** OfficeKit rejects the operation before reading or writing bytes

### Requirement: Rights policy fails closed
OfficeKit SHALL allow only Public Domain, CC0, CC BY, Lucide ISC, explicit user/generated/permission declarations, and official press-kit declarations; all unknown, ShareAlike, NonCommercial, and NoDerivatives assets SHALL be rejected.

#### Scenario: CC BY metadata is complete
- **WHEN** a CC BY candidate has author and license URL metadata
- **THEN** OfficeKit stores a credit line and marks visible attribution as required

#### Scenario: CC BY metadata is incomplete
- **WHEN** a CC BY asset lacks its author or license URL
- **THEN** OfficeKit rejects acquisition instead of inventing attribution

#### Scenario: Openverse evidence boundary
- **WHEN** a candidate came from Openverse
- **THEN** its receipt labels the rights evidence as provider-declared and does not claim legal verification

### Requirement: Remote downloads are bounded and destination-safe
OfficeKit SHALL accept only HTTPS URLs without credentials, reject unsafe destinations before connecting, pin a validated public address for each request, revalidate every redirect, and enforce byte, image dimension, MIME, and magic-byte limits.

#### Scenario: Private destination
- **WHEN** a hostname resolves to loopback, private, link-local, or cloud metadata space
- **THEN** download is rejected before an HTTP request is sent

#### Scenario: Redirect changes trust boundary
- **WHEN** an allowed public URL redirects toward an unsafe host or exceeds three redirects
- **THEN** download fails without following the unsafe or excess hop

#### Scenario: Oversized or mismatched content
- **WHEN** a response exceeds 20 MiB, 40 megapixels, 16,384 pixels on one edge, or its MIME disagrees with image magic
- **THEN** no task asset or receipt is published

#### Scenario: Unsupported remote format
- **WHEN** a remote response is WebP, SVG, or a non-image payload
- **THEN** acquisition fails without transcoding

### Requirement: Offline Lucide icons are deterministic
OfficeKit SHALL search the pinned Lucide collection locally and materialize a selected icon as a safe SVG with ISC provenance.

#### Scenario: Icon search and add
- **WHEN** an English icon query matches Lucide names
- **THEN** stable candidates are returned without network access and the selected candidate produces deterministic SVG bytes and an ISC receipt
