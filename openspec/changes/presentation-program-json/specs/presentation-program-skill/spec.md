## ADDED Requirements

### Requirement: PPJ-first progressive Presentation Skill
The main Presentations Skill SHALL be a short router for create, import, edit, continue, review, and deliver workflows and SHALL identify PPJ as the only public authored presentation source.

#### Scenario: Fresh Agent creates a presentation
- **WHEN** a fresh Agent receives an ordinary presentation request
- **THEN** the Skill routes it to PPJ guidance without loading legacy MJS/Compose examples or unrelated advanced references

### Requirement: Complete PPJ language reference
The packaged Skill SHALL include one generated `ppj.md` reference derived from the schema and capability registry, with typed fields, limits, minimal examples, and path-specific error guidance.

#### Scenario: New typed primitive is documented
- **WHEN** a public PPJ field or native capability is added
- **THEN** the generated reference and capability registry identify its syntax, owner, review requirement, and minimum example

### Requirement: Focused presentation references
The Skill SHALL route on demand to focused references for fonts, shapes, text, charts/tables, media/layers, motion, components/templates, imported native references, scenario design, and review/delivery.

#### Scenario: Imported SmartArt edit loads bounded guidance
- **WHEN** a task needs an imported SmartArt text edit
- **THEN** the Agent loads the PPJ and imported-native reference without loading unrelated from-scratch design or motion catalogs

### Requirement: Presentation quality invariants
The Skill SHALL prohibit fabricated evidence, card-wall defaults, random geometry, meaningless decoration, accidental information occlusion, and unrelated multi-accent AI palettes while preserving user templates and explicit design authority.

#### Scenario: Chart label is occluded
- **WHEN** a rendered data line, bar, label, or marker obstructs another required value
- **THEN** review requires a truthful layout or styling correction rather than hiding the data relationship

### Requirement: Capability registry has complete ownership
Every stable public Presentation capability SHALL be assigned to PPJ state, nativeRef, compiler/helper, inspect/review, or host-only, with exactly one canonical Skill/Help owner.

#### Scenario: Orphan capability enters Help
- **WHEN** Help contains a new public Presentation capability without a registry owner or PPJ classification
- **THEN** the Skill consistency gate fails with the missing capability name

### Requirement: Presentation Skill Maintainer
The package SHALL include a host-neutral Presentation Skill Maintainer that explains the required schema, compiler, Help, generated reference, review, example, and registry updates for a primitive change.

#### Scenario: Maintainer handles a new chart field
- **WHEN** an Agent uses the Maintainer for a new chart capability
- **THEN** it identifies every required contract owner without rewriting unrelated design guidance

### Requirement: Legacy public authoring guidance is removed
After PPJ acceptance, the package SHALL remove Presentation/MJS/Compose from the public Skill route and public package entrypoint.

#### Scenario: Packed clean install
- **WHEN** a fresh Agent inspects the OfficeKit 2.0 package
- **THEN** PPJ is the only public file-authoring route and no packaged task document presents MJS/Compose as an equivalent default
