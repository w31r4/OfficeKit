## Why

F-07 chart typography now preserves outer shadows but cannot express glow or retain a known glow-plus-shadow character effect list. Ordinary PPJ text already defines the reusable color/radius/opacity field.

## What Changes

- Add `chartTextStyle.glow` using the ordinary `glow` schema: required color and radius, optional opacity.
- Preserve RGB/theme identity, grammar color and opacity tokens, explicit zero radius/alpha and alpha absence across native chart text, projection and source-bound edits.
- Support glow alone or followed by the existing outer shadow, including independent deletion of either effect.
- Carry glow through style precedence and vector text defaults/overrides; add minimal lifecycle and preservation evidence.

## Capabilities

### New Capabilities
- `ppj-chart-text-glow`: Direct chart text glow and coexistence with outer shadow.

### Modified Capabilities

None. No main specs are present; this increment expands the chart effects profile established by the earlier shadow change.

## Impact

PPJ schema, additive wire field, shared DrawingML glow/effects codecs, chart authoring/projection/edit/vector paths, JS wire preservation, field reference and coverage. Full F-07, other effect graphs and host visual acceptance remain open.
