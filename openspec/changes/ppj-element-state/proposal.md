## Why

PPJ v1 declares optional `hidden` and `locked` state on every presentation
element, and the generated Agent reference advertises those fields. The
authored compiler currently rejects both unconditionally. This makes the
language contract false and prevents an Agent from protecting guides,
background layers, media posters, or finished visual groups without leaving
PPJ.

## What Changes

- Define one cross-type Agent meaning for `hidden` and `locked`.
- Add optional element state to the additive Presentation wire model.
- Compile authored state into canonical PresentationML non-visual properties
  and type-appropriate DrawingML lock profiles.
- Project recognized third-party state into PPJ and issue a source-bound
  capability only for an exact canonical locked or unlocked profile.
- Preserve partial, private, or otherwise unrecognized native lock profiles
  without pretending that a boolean can edit them safely.
- Keep Office-required baseline locks, such as chart/table `noGrouping` and
  media-picture `noChangeAspect`, separate from PPJ `locked`.
- Synchronize the generated language reference, shape/layer guidance,
  capability registry, and coverage evidence.

## Capabilities

### New Capabilities

- `ppj-element-state`: Authored and safely source-bound element visibility and
  canonical edit locking across ordinary presentation object kinds.

### Modified Capabilities

None. The schema ID and Office wire protocol version remain unchanged.

## Impact

- `proto/office_kit/artifact/v1/office_artifact.proto` receives additive
  optional state and source-capability fields.
- PPJ parsing, authored lowering, PPTX projection/export, semantic hashing,
  generated PPJ documentation, capability ownership, and one existing PPJ
  contract test are affected.
- No raw OOXML surface, JavaScript authoring API, host automation, or universal
  interpretation of arbitrary native lock combinations is introduced.

