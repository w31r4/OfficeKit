## ADDED Requirements

### Requirement: One Presentation Template schema with optional references
Presentation Template schema v3 SHALL retain `SKILL.md`, preview, and representative examples as required fields and SHALL permit optional, declared, hashed, and licensed `referenceProgram` and `referencePptx` assets.

#### Scenario: Style-only template remains valid
- **WHEN** a v3 template contains the required guide and images but no reference program or PPTX
- **THEN** search and selection continue to return the same style evidence without a second template kind

#### Scenario: Reference-backed template is valid
- **WHEN** a v3 template declares matching local PPJ/PPTX files, hashes, rights, and provenance
- **THEN** search returns those references as optional evidence without making page cloning the default workflow

### Requirement: Template Creator builds executable calibration
The Presentation Template Creator SHALL create an original PPJ and compiled PPTX during calibration, verify them, and include them in the published template only when the chosen rights and package policy allow publication.

#### Scenario: Reference cannot be published
- **WHEN** the calibration source or generated reference lacks sufficient publication rights
- **THEN** the Creator retains the program, deck, and analysis as Task evidence and publishes only the clean-room guide and original permitted examples

### Requirement: Template consumption remains style-first
The Agent SHALL first read the template guide and representative images, form the current deck's Design Grammar, and use optional reference PPJ/PPTX only for exact components, assets, or native structures that the current task needs.

#### Scenario: New topic uses a reference-backed template
- **WHEN** an Agent creates a deck from unrelated content with a selected template
- **THEN** it composes new pages for the communication task and does not clone the whole reference deck by default

### Requirement: Evidence Ledger clean-room template
The package SHALL add a distinct Evidence Ledger template with original content and geometry, complete reference PPJ/PPTX, and representative evidence for cover, hypothesis, method, timeline, quantitative result, confidence interval, and decision pages.

#### Scenario: Evidence Ledger package validates
- **WHEN** the template is scanned in the release package
- **THEN** its guide, preview, examples, reference program, reference deck, hashes, provenance, and license declarations all match packaged bytes

### Requirement: Existing templates and WIP remain independent
The Evidence Ledger work SHALL NOT alter Cranberry Evidence or consume uncommitted shared-worktree template files.

#### Scenario: Template change is reviewed
- **WHEN** the Evidence Ledger commit is inspected
- **THEN** no Cranberry Evidence path or unrelated user WIP appears in its diff
