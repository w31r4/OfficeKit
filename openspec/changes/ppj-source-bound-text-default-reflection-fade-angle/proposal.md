## Why

Paragraph default-text reflections already preserve the full-span `fadeDir`
token in the native codec, but PPJ does not expose a native field for it. A
caller cannot make a narrow source-bound edit to the reflection fade direction
without treating the effect as source-owned.

## What Changes

- expose `text.paragraphs[].style.defaultText.reflection.fadeAngle` as
  `textDefaultReflectionFadeAngleDegrees` for a strict direct paragraph
  default reflection;
- validate the native 1/60000-degree token in the existing unsigned
  `0 <= fadeDir < 21600000` range and preserve the raw integer;
- patch only `a:reflection/@fadeDir` in the owning slide while preserving all
  other reflection attributes, sibling effects, text topology, and unrelated
  ZIP parts;
- reject missing or malformed fade tokens, bounds errors, unsupported effect
  graphs, stale authority, and topology or identity changes;
- add a focused imported/authored source-bound lifecycle experiment and update
  the PPJ schema, capability registry, presentation reference, coverage, and
  F-03 backlog evidence.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-default-reflection-fade-angle`: source-preserving
  editing of one paragraph default-text reflection fade-direction token.

### Modified Capabilities

<!-- No existing OpenSpec capability requirements are present in this repo. -->

## Impact

The native PPJ leaf projection and source-bound edit-plan proof gain one
paragraph default-reflection scalar. The public PPJ contract and generated
presentation Skill gain the matching field/capability; no protobuf wire change
is required. Other reflection transforms, ordinary run effects, and complex
effect graphs retain their existing boundaries.
