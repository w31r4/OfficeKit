# PPJ master and layout authoring

## ADDED Requirements

### Requirement: PPJ authors bounded native Master/Layout state

PPJ SHALL express one source-free slide master, bounded layouts, direct
backgrounds, master paragraph defaults, owner-local text placeholders, and page
layout bindings as typed persistent state.

#### Scenario: Authored layout-bound presentation

- **WHEN** a valid source-free PPJ declares one master, one or more supported
  layouts, and pages bound to those layouts
- **THEN** NativeAOT SHALL compile native editable Master and Layout parts
- **AND** a second import SHALL recover each page's layout binding.

#### Scenario: Invalid master or placeholder graph

- **WHEN** a source-free PPJ declares multiple masters, an unsupported layout
  type, an unsupported placeholder type, a duplicate text-style level, or a
  missing master/layout reference
- **THEN** validation SHALL reject it before output changes.

### Requirement: Imported layout identity remains source-owned

PPJ SHALL expose the stable layout identity of a projected third-party page
without granting authority to rebuild or mutate the source Master/Layout graph.

#### Scenario: Third-party layout binding is unchanged

- **WHEN** a third-party PPTX is projected to PPJ and rebuilt without changing
  its projected page layout values
- **THEN** the original source package SHALL be returned byte-for-byte.

#### Scenario: Third-party layout binding is changed

- **WHEN** an Agent changes a projected source page's layout value without an
  issued capability
- **THEN** the build SHALL fail closed without rebuilding the Master/Layout
  graph.

### Requirement: PPJ-state registry entries identify real language paths

The capability registry SHALL map every Help API classified as persistent PPJ
state to one concrete path in the public language schema.

#### Scenario: A PPJ-state API has no language mapping

- **WHEN** a Help API is classified as `ppj-state` but has no concrete PPJ path
- **THEN** the Presentation Skill maintainer SHALL fail.
