## ADDED Requirements

### Requirement: Authored structured chart title
The NativeAOT compiler SHALL accept PPJ chart `title` in either string or
structured `textContent` form and SHALL lower structured paragraphs and runs
to native ChartPart rich text.

#### Scenario: Two-tone analytical title
- **WHEN** a valid authored chart title contains two runs with distinct weight,
  color, and East Asian font choices
- **THEN** build writes one editable native chart title whose visible text and
  run formatting match the PPJ while retaining chart data and frame identity

### Requirement: Uniform style defaulting
Chart `titleTextStyle` SHALL act as a default for structured-title run
properties and SHALL NOT override an explicit run property.

#### Scenario: Accent run overrides title default
- **WHEN** `titleTextStyle` declares a dark default color and one title run
  declares an accent color
- **THEN** all unstyled runs use the default and the explicit accent run keeps
  its own color after build and reimport

### Requirement: Conservative imported projection
The PPTX projector SHALL recover a structured chart title only when its native
rich-text container is fully inside the bounded paragraph/run profile.

#### Scenario: Formula-backed title
- **WHEN** a third-party chart title is driven by a formula or contains an
  unrecognized rich-text extension
- **THEN** OfficeKit preserves the source title but does not claim a typed
  structured-title edit capability

### Requirement: Source-bound local title edit
A source-bound structured title edit SHALL require a fresh chart capability
and SHALL change only the owned title rich-text subtree plus unavoidable XML
namespace serialization.

#### Scenario: Restyle one imported title run
- **WHEN** an Agent changes one bounded run's text and color in an imported
  chart title
- **THEN** build preserves chart series, axes, relationships, embedded data,
  and every non-target OPC part while reimport recovers the new run semantics

### Requirement: Agent discoverability
The generated PPJ reference and chart guidance SHALL show when to use a string
title, when structured runs add information, and which imported title profiles
remain source-bound.

#### Scenario: Agent emphasizes the measured delta
- **WHEN** an Agent searches the PPJ guidance for a partially emphasized chart
  title
- **THEN** it can find a minimal structured-title example without reading the
  C# codec or raw ChartML
