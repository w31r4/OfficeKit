## Why

OfficeKit already proves and performs bounded source-component reuse: it
clones a complete source slide, removes only codec-proven siblings, checks the
retained native identity and connector graph, and preserves unknown descendants
inside the retained top-level object. PPJ can reuse a complete source page but
cannot yet express this higher-value template-continuation state.

## What Changes

- Extend PPJ `sourceClone` with one optional `retainElement` ID.
- Interpret an absent `retainElement` as the existing exact full-page clone.
- Interpret a present `retainElement` as a finite component projection that
  keeps one exact top-level source element and deletes every sibling.
- Require the source page's issued `duplicate/pageClone` capability and every
  omitted sibling's independently issued `delete/element` capability.
- Lower the request through the existing native clone and element-deletion
  writer, then require build/reimport before editing the retained component.
- Teach Agents to use the result as a source-derived page, not as a general
  element-copy or arbitrary graph-extraction primitive.

## Capabilities

### New Capabilities

- `ppj-source-component-reuse-parity`: Capability-issued extraction of one
  exact top-level source component into a new slide.

### Modified Capabilities

- `ppj-source-slide-reuse-parity`: `sourceClone` gains the optional bounded
  `retainElement` selector while preserving existing full-page semantics.

The PPJ schema ID and Office wire protocol version remain unchanged.

## Impact

PPJ schema/model validation, source-bound lowering, generated Skill guidance,
coverage, and one existing focused C# contract are affected. The native clone
codec, element-deletion proof, Office wire, JavaScript candidate workflow, and
authored compiler remain unchanged.
