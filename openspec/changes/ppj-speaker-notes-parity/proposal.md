## Why

OfficeKit already authors, imports, validates, adds, and locally edits a bounded
speaker-notes profile. PPJ authored decks can contain notes, but third-party
projection currently flattens structured paragraphs and runs to one string, and
the source-bound PPJ compiler rejects every notes change. The public language
therefore loses semantic state before it reaches an existing capable codec.

## What Changes

- Project supported imported notes as PPJ `textContent`, preserving paragraph,
  run, and supported formatting state.
- Issue a page-level `setNotes` capability only when the native codec proves an
  existing notes body editable or an absent notes body safely addable.
- Lower capable PPJ text-only notes changes through the existing source-bound
  speaker-notes codec.
- Preserve notes-master, layout, fields, hyperlinks, picture bullets, unknown
  relationships, and unsupported body topology as source-owned state.
- Synchronize the generated PPJ reference, text/review guidance, coverage, and
  one existing comprehensive PPJ contract.

## Capabilities

### New Capabilities

- `ppj-speaker-notes-parity`: Structured projection and bounded source-bound
  editing of the existing Presentation speaker-notes profile.

### Modified Capabilities

None. The PPJ schema ID and Office wire protocol version remain unchanged.

## Impact

- PPJ schema/model parsing, PPTX projection, source-bound diff lowering,
  generated documentation, and one existing PPJ contract are affected.
- The protobuf and native NotesSlide codec already carry the required state, so
  no wire or OOXML writer change is required.
- No notes-master editing, arbitrary notes shape editing, raw OOXML, media,
  hyperlinks, fields, or comment behavior is introduced.
