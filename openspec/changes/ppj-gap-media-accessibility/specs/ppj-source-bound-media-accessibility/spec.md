## ADDED Requirements

### Requirement: Canonical imported media exposes its non-visual owner

When an imported audio/video picture has a valid residual `p:cNvPr` profile,
the projector MUST keep the element `type` as `opaque` with
`nativeKind: "media"`, MAY expose its common `accessibility` object, and MUST
issue `setAccessibility` for that element. The media relationships, poster,
and playback graph remain opaque-owned.

#### Scenario: Imported media projects alternative text capability

- **WHEN** a source slide contains a media picture with an unambiguous
  residual `cNvPr` accessibility profile
- **THEN** source-bound PPJ contains the title/description/decorative state
  and a `setAccessibility` capability while retaining `nativeKind: "media"`

### Requirement: Source-bound media metadata edits stay local

For a capability-bound imported media element, a source-bound PPJ compile MAY
add, replace, or clear only `accessibility.title`,
`accessibility.description`, and `accessibility.decorative`. The native writer
MUST update only the media picture's non-visual owner, preserve its other
children and attributes, and report the owning SlidePart as changed.

#### Scenario: Media accessibility round trips without touching playback

- **WHEN** source-bound PPJ adds a title and description to an imported media
  element
- **THEN** the output validates, keeps the media relationship IDs and native
  playback markers, and a second projection recovers the requested metadata

### Requirement: Ambiguous media owners fail closed

If the media `cNvPr` residual profile contains an ambiguous known decorative
graph or invalid modeled metadata, the importer MUST NOT issue
`setAccessibility`, and a semantic accessibility mutation MUST fail closed
without rebuilding or flattening the media object.

#### Scenario: Irregular media metadata stays opaque

- **WHEN** an imported media picture has an unsupported or ambiguous
  non-visual accessibility graph
- **THEN** the source remains byte-preserved and the request is rejected as an
  unsupported presentation edit
