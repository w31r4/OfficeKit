## Why

F-03 still lacks complete language/script layout semantics. PPJ can express column order and vertical text, but cannot author or independently edit a paragraph's native right-to-left direction; treating these other fields as substitutes loses meaning.

## What Changes

- Add paragraph `direction: "left-to-right" | "right-to-left"`, preserving explicit LTR, RTL and absence separately. Direction applies independently of alignment, column order and vertical text.
- Add an optional native wire boolean using a new field number; retain existing wire version and fields. Generate bindings with the repository script.
- Support authored style precedence, native import/reprojection, and ordinary source-bound text/shape add/set/remove/restore under exact direction authority. Deleting the field or its direction-only style wrapper clears the direct native attribute.
- Preserve equivalent native boolean spelling, neighboring paragraphs, runs and unrelated XML/ZIP. Unknown native direction tokens stay source-owned and reject replacement while independent edits remain possible.
- Synchronize schema, Help, registry, generated references, Agent guidance and the F-03 backlog. Preview explicitly reports unresolved bidirectional layout.

## Capabilities

### New Capabilities

- `ppj-paragraph-direction-lifecycle`: explicit paragraph direction and its source-preserving lifecycle.

### Modified Capabilities

None. There is no corresponding main spec.

## Impact

PPJ schema and compilers/projector, PresentationTextParagraph wire and generated binding, shared paragraph codec, Help and presentation documentation, focused native and preview tests. No new dependency, inherited-direction solver, Unicode bidi/shaping engine or host layout acceptance is introduced.
