## Why

Paragraph default-text shadows already round-trip authored `scaleY`, but an
imported default-run shadow still lacks a native field for its direct vertical
scale token. A caller therefore cannot make a narrow source-bound edit without
treating the whole effect graph as source-owned; this closes the next F-03
default-text effect leaf beside `scaleX`.

## What Changes

- expose `text.paragraphs[].style.defaultText.shadow.scaleY` as a source-bound
  native field for a strict direct paragraph default outer shadow;
- validate the existing signed scale range and native 1/100000 precision;
- patch only `a:outerShdw/@sy` in the owning slide while preserving all sibling
  effects, color/opacity, other shadow transforms, paragraph/run topology, and
  unrelated ZIP parts;
- reject missing or malformed transforms, unsupported effect graphs, stale
  authority, and combined topology or identity changes;
- add a focused imported/authored round-trip experiment and update the PPJ
  schema, capability registry, presentation reference, coverage, and backlog.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-default-shadow-scale-y`: source-preserving editing of
  one paragraph default-text outer-shadow vertical scale token.

### Modified Capabilities

<!-- No existing OpenSpec capability requirements are present in this repo. -->

## Impact

The native PPJ leaf projection and source-bound edit-plan proof gain one
paragraph default-shadow scalar. The public PPJ contract and generated
presentation Skill gain the matching field/capability; no protobuf wire change
is required. Other shadow transforms, automatic effects, and complex effect
graphs retain their existing boundaries.
