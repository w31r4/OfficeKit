## ADDED Requirements

### Requirement: Presentation images accept FileBlob bytes
Presentation image placement and Compose image nodes SHALL accept `blob: FileBlob` in addition to existing image inputs, validate supported image MIME, and normalize the bytes into the existing serialized image representation.

#### Scenario: Direct image placement
- **WHEN** an Agent loads a registered PNG asset as a `FileBlob` and passes it to `slide.images.add`
- **THEN** the exported PPTX embeds the identical PNG bytes with the requested placement, fit, crop, and accessibility description

#### Scenario: Ambiguous byte source
- **WHEN** both `blob` and `dataUrl` are provided for one image
- **THEN** image creation fails rather than choosing one silently

#### Scenario: Compose image placement
- **WHEN** a Compose image node carries a valid `FileBlob`
- **THEN** materialization uses the same canonical image path without requiring caller-built base64

### Requirement: Presentations route image sourcing progressively
The Presentations Skill SHALL route to one host-neutral image-sourcing reference only when a page has a declared media role or an existing image must be replaced.

#### Scenario: External image is useful
- **WHEN** a page needs photographic evidence, identity, explanation, or atmosphere not supplied by the user or template
- **THEN** the Agent searches in English, inspects a small candidate set, selects an allowed asset, adds it to the task, and embeds the returned local file

#### Scenario: No compliant image is useful
- **WHEN** no candidate serves the communication task or can be reviewed safely
- **THEN** the Agent uses a native visual or no image and does not add decorative stock merely to fill space

#### Scenario: Visual understanding is unavailable
- **WHEN** the Agent cannot visually review imagery
- **THEN** it may use deterministic Lucide icons, but it does not claim a decorative photograph was visually approved and instead requests human review or chooses a native visual

### Requirement: Delivery review covers imagery and rights
Presentation review and delivery guidance SHALL require image crop, contrast, clarity, repetition, alternative text, source display, attribution, and audit evidence appropriate to each image role.

#### Scenario: Evidence image delivery
- **WHEN** a sourced image is used as evidence
- **THEN** the rendered page carries an appropriate visible source and the audit report links the embedded bytes to the task receipt

#### Scenario: Decorative CC BY image delivery
- **WHEN** a decorative CC BY image is used
- **THEN** its required credit appears visibly on the page or a credits page and the sources sidecar records the same obligation

#### Scenario: Existing source-bound image replacement
- **WHEN** an imported PPTX image is replaced with a registered task asset
- **THEN** the existing source-bound replacement path performs the edit and the new asset remains linked to its provenance receipt

