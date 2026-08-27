## ADDED Requirements

### Requirement: The specialist creator distills PowerPoint style
The `presentation-template-creator` Skill SHALL accept a reference deck, visual
references, written direction, or an OfficeKit task and guide the Agent to
extract only reusable style decisions.

#### Scenario: Reference PPTX supplied
- **WHEN** the user asks to save a presentation's style as a template
- **THEN** the Creator renders and inspects it as reference evidence, writes an
  independent style guide, and does not copy the source file into the result

### Requirement: Examples are original calibration work
The Creator SHALL produce four to six unrelated OfficeKit-authored calibration
pages spanning at least three declared roles and SHALL package only their PNG
renders as visual examples.

#### Scenario: Built-in template migration
- **WHEN** an old bundled template is rebuilt
- **THEN** its new examples use new content, geometry, and wording and contain
  no old screenshot or retained source page

### Requirement: Packaging is deterministic and bounded
The Creator's packaging script SHALL validate the fixed directory surface,
English search metadata, PNG dimensions, example roles, hashes, license fields,
and safe relative paths before atomically publishing a Template Skill.

#### Scenario: Successful package
- **WHEN** a complete guide, metadata draft, and valid example images are passed
  to the script
- **THEN** it writes the Skill, generated montage, Agent metadata, and canonical
  schema-v3 sidecar with content hashes

#### Scenario: Existing template update is stale
- **WHEN** an update's expected current hash does not match the target Skill
- **THEN** the script refuses to overwrite it and leaves the existing Skill
  unchanged

### Requirement: Source evidence stays outside the template
Reference files, intermediate presentations, extracted media, and creator QA
evidence MUST remain in the task workspace and MUST NOT be copied into the
published Template Skill.

#### Scenario: Package surface is audited
- **WHEN** a generated template directory is scanned
- **THEN** it contains only the fixed guide, JSON, Agent YAML, and PNG asset
  surface

### Requirement: The specialist is installed and routed by default
`officekit init` SHALL install `presentation-template-creator`, the OfficeKit
router SHALL select it for PowerPoint template authoring, and the generic
creator SHALL remain responsible only for DOCX/XLSX templates.

#### Scenario: Fresh initialization
- **WHEN** OfficeKit initializes a supported Agent workspace
- **THEN** both the generic creator and presentation specialist are available
  and a PPT template request routes to the specialist
