## Why

F-07 chart typography can now retain glow and outer shadow but still cannot express the ordinary PPJ soft-edge field. A soft edge also makes an otherwise known chart character effect list source-owned.

## What Changes

- Add `chartTextStyle.softEdge` using `{ radius }`, matching the ordinary 0–1000 pt field and native EMU precision.
- Preserve explicit zero separately from omission across chart styles, trendline rich styles, source-bound edits and fresh projection.
- Compose soft edge after known glow/outer shadow and permit independent removal/restoration of all modeled effects.
- Align authored shadow/glow resource-reference validation with the already supported direct theme-color contract, retaining wrong-kind and RGB-only field guards.
- Propagate field precedence and vector defaults/overrides; add focused lifecycle and native graph regression evidence.

## Capabilities

### New Capabilities
- `ppj-chart-text-soft-edge`: Chart character soft edge with independent effect composition and source-bound lifecycle.

### Modified Capabilities

None. Main specs are empty; this increment extends the earlier chart glow/shadow effect profile.

## Impact

Schema and additive wire field, shared direct soft-edge validation, chart effects adapter and authored/source-bound/projector/vector mappings, JS preservation and synchronized field references. Full ChartML and host appearance remain open.
