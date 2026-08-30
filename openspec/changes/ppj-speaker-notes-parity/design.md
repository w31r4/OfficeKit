## Context

`PptxSpeakerNotesCodec` already owns a conservative, hash-bound notes profile.
It preserves a relationship-free body placeholder, supports plain text and
bounded paragraph/run formatting, can add one canonical notes leaf when the
presentation graph permits it, and rejects irregular notes topology. PPJ only
uses that profile for authored builds today.

## Goals / Non-Goals

**Goals:**

- Preserve representable paragraph/run state when importing notes to PPJ.
- Let an Agent edit note text through the same page-level capability model used
  for backgrounds and transitions.
- Keep formatting and run/paragraph topology fixed during source-bound edits.
- Reuse the native codec's source hash and graph proof.

**Non-Goals:**

- Editing NotesMaster, notes layout, notes-local shapes, fields, hyperlinks,
  picture bullets, media, relationships, or arbitrary XML.
- Adding or removing paragraph/run topology in an imported rich notes body.
- Deleting an imported NotesSlide or converting plain and rich representations.

## Decisions

### 1. Notes remain ordinary PPJ text content

`pages[].notes` keeps its existing `textContent` shape. A simple notes body is
projected as a string; a semantically rich but supported body is projected as
paragraphs and runs. PPJ does not gain a second notes-specific text schema.

### 2. Mutation is one page capability

Projection issues `setNotes` with field `notes` when either the existing notes
source binding is editable or the slide source binding says a canonical notes
leaf is addable. Capability issuance is evidence, not authority; build still
re-proves all source hashes and package topology.

### 3. Imported edits preserve topology and style

For an existing rich body, only run text values may change. Paragraph IDs, run
IDs, style objects, counts, and representation kind must remain identical. For
an absent notes part, the bounded add profile accepts plain text only.

### 4. Native source binding stays out of PPJ

The notes part path, relationship ID, XML hash, and semantic hash remain in the
Presentation wire object recovered from the fresh source projection. PPJ carries
only the page native capability and semantic notes value, keeping package-local
identities out of Agent context.

## Risks / Trade-offs

- [PPJ rich text accepts more formatting than imported notes can mutate] ->
  Compare the complete non-text JSON structure and reject style/topology edits.
- [A caller forges a capability] -> Require the page capability hash to match
  the fresh projection, then rely on `PptxSpeakerNotesCodec` to re-prove source
  binding and graph state.
- [An empty existing NotesSlide looks absent in PPJ] -> Use the fresh wire source
  object to retain its binding when text is added.

## Migration Plan

The schema expansion is additive. Existing authored PPJ is unchanged. Imported
PPJ gains richer notes values and, only where proven, one additional native
capability.

## Open Questions

None.
