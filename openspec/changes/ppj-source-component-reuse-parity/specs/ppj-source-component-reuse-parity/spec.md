## ADDED Requirements

### Requirement: Declarative component clone
A source-bound PPJ SHALL express one pending exact top-level component clone by
adding `retainElement` to the existing finite `sourceClone` descriptor.

#### Scenario: Agent retains one source component
- **WHEN** an Agent inserts an empty pending clone immediately after its source
  page and names one top-level source element
- **THEN** build clones the proven source graph and removes every other direct
  source element through existing native deletion proofs

### Requirement: Composed capability authority
Component reuse SHALL require the source page's fresh `duplicate/pageClone`
capability and each omitted sibling's fresh `delete/element` capability.

#### Scenario: One sibling is not independently deletable
- **WHEN** any omitted source element lacks a supported deletion capability
- **THEN** check or build rejects the component clone without changing the PPTX

### Requirement: Retained graph remains exact
The pending component clone SHALL keep the selected source element and its
owned descendants unchanged until build/reimport.

#### Scenario: Agent selects a nested child or edits pending state
- **WHEN** `retainElement` does not identify a direct page element, or the
  pending page declares native content or editable source state
- **THEN** check or build rejects the program instead of flattening or guessing

### Requirement: Native graph is independently re-proven
The native writer SHALL revalidate the actual source binding, shape-tree order,
native IDs, relationships, connector targets, and deletion closure.

#### Scenario: Semantic IDs no longer match the native graph
- **WHEN** the source bytes, capability set, or deletion topology changed since
  PPJ projection
- **THEN** build fails closed before publishing an output package

### Requirement: Reimport yields ordinary editable state
After component reuse is built, PPJ projection SHALL return an ordinary
source-bound page containing the retained component and no pending macro.

#### Scenario: Agent continues after reuse
- **WHEN** the output PPTX is projected back to PPJ
- **THEN** the new page has no `sourceClone`, exposes the retained top-level
  object with nativeRef, and can use its independently issued edit capabilities

### Requirement: Existing full-page reuse remains compatible
Omitting `retainElement` SHALL preserve the existing exact full-slide clone.

#### Scenario: Old PPJ uses sourceClone
- **WHEN** a valid source-bound PPJ contains only `page` and `capability`
- **THEN** it builds with unchanged full-page clone semantics
