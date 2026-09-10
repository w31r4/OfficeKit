## Why

Paragraph default-text shadows now expose imported `skewX`, but the matching
vertical `skewY` token still lacks a native field. A caller cannot make a
narrow source-bound edit without treating the whole effect graph as
source-owned; this completes the paired default-text skew leaves for F-03.

## What Changes

- expose `text.paragraphs[].style.defaultText.shadow.skewY` as a source-bound
  native field for a strict direct paragraph default outer shadow;
- validate the strict `-90 < skewY < 90` range after native 1/60000-degree
  rounding and preserve the signed Int32 token;
- patch only `a:outerShdw/@ky` in the owning slide while preserving all sibling
  effects, color/opacity, other shadow transforms, paragraph/run topology, and
  unrelated ZIP parts;
- reject missing or malformed transforms, unsupported effect graphs, stale
  authority, and combined topology or identity changes;
- add a focused imported/authored round-trip experiment and update the PPJ
  schema, capability registry, presentation reference, coverage, and backlog.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-default-shadow-skew-y`: source-preserving editing of
  one paragraph default-text outer-shadow vertical skew token.

### Modified Capabilities

<!-- No existing OpenSpec capability requirements are present in this repo. -->

## Impact

The native PPJ leaf projection and source-bound edit-plan proof gain one
paragraph default-shadow scalar. The public PPJ contract and generated
presentation Skill gain the matching field/capability; no protobuf wire change
is required. Other shadow transforms, automatic effects, and complex effect
graphs retain their existing boundaries.
