## ADDED Requirements

### Requirement: Authored media declares a bounded playback trigger

An authored PPJ `media` element MAY contain `playback.trigger`. The trigger
MUST be either `onClick` or `onSlideStart`; omission MUST retain the existing
click-start behavior.

#### Scenario: Slide-start playback lowers to the media timing owner

- **WHEN** an authored video declares `playback.trigger: "onSlideStart"`
- **THEN** the compiler emits a valid media timing node whose start condition
  is immediate, while retaining the existing media payload, poster, and
  relationship closure

#### Scenario: Omitted playback remains compatible

- **WHEN** an authored media element omits `playback`
- **THEN** the compiler emits the previous click-start timing condition

### Requirement: Trigger declaration survives authored recovery

When a source-free authored PPTX is projected through its embedded PPJ
snapshot, the recovered program MUST retain the declared playback trigger.

#### Scenario: Embedded PPJ recovers slide-start intent

- **WHEN** a compiled PPTX containing an authored `onSlideStart` media trigger
  is projected through the embedded-program path
- **THEN** the recovered PPJ contains the same `playback.trigger` value

### Requirement: Imported timing remains source-owned

The authored playback trigger profile MUST NOT grant edit permission for an
imported or third-party media timing graph. Such media remains opaque and any
unsupported trigger/timing mutation MUST fail closed.

#### Scenario: Complex imported media timing is not rebuilt

- **WHEN** a source-bound PPTX contains a media timing graph outside the
  authored profile
- **THEN** projection does not pretend the graph is an authored media object
  and source-bound compilation does not replace it with a guessed trigger
