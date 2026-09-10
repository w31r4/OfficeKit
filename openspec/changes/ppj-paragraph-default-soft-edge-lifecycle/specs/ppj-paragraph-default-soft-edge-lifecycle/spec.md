## Purpose

Complete direct paragraph default soft-edge editing in PPJ while preserving the surrounding imported presentation and native effect state.

## ADDED Requirements

### Requirement: Independent paragraph default soft edge

Ordinary text/shape paragraphs SHALL support `text.paragraphs[].style.defaultText.softEdge` assignment, removal and restoration under that exact source-bound field authority. The object SHALL require only a finite numeric `radius` in 0..1000 points, rounded to the nearest EMU with ties to even. Explicit zero SHALL retain a zero-radius effect; removing the field or its single-effect wrapper SHALL remove the direct effect. Missing radius, invalid values, token objects and unknown fields SHALL reject without output.

#### Scenario: Zero, fractional radius and removal
- **WHEN** a paragraph soft edge is assigned zero, a fractional radius or the upper endpoint, removed, then restored
- **THEN** exported native state and fresh PPJ SHALL preserve the rounded value or absence independently, including wrapper removal
- **AND** an edit lacking the exact field authority SHALL reject without output

### Requirement: Source preservation

Soft-edge edits SHALL preserve other effect nodes and list attributes, direct runs, neighboring paragraphs and non-target XML/ZIP content. Unknown attributes/children, missing or invalid radius, duplicate effect/list nodes and effect DAGs SHALL remain source-owned: no-op preserves the source bytes, unrelated scalar edits preserve their XML, and replacement rejects without output.

#### Scenario: Mixed or unmodeled source effects
- **WHEN** an imported paragraph has a mixed effect list or unmodeled soft-edge state
- **THEN** bounded edits SHALL change only their target state and preserve other source content
- **AND** unsupported soft-edge replacement SHALL reject without output

### Requirement: Discoverability and preview evidence

Schema authority, Help, registry, generated references and focused text guidance SHALL describe this lifecycle consistently. Preview assessment SHALL report partial support rather than silently discarding the soft edge or claiming host fidelity.

#### Scenario: Preview of a paragraph soft edge
- **WHEN** a paragraph default contains a soft edge, including radius zero
- **THEN** preview assessment SHALL include an explicit support-limit diagnostic for that field
