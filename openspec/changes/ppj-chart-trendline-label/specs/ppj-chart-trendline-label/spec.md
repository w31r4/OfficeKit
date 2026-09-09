## Purpose

Represent native trendline labels as persistent PPJ state that can be authored, inspected, edited and removed without losing the surrounding chart.

## ADDED Requirements

### Requirement: Bounded trendline label state
`trendlines[].label` SHALL accept optional `text`, `numberFormat`, `textStyle`, `fill` and `line` using the existing bounded chart-label grammar. Text and number formats SHALL resolve to nonempty strings of at most 255 characters without controls. Text SHALL use a single literal native run. Omitting text SHALL retain automatic equation/R-squared content; `{}` SHALL preserve the presence of the default label container. Label edits SHALL preserve the trendline display flags.

#### Scenario: Author a styled equation label
- **WHEN** an ordinary bar/column/line or categorical combo trendline declares a label
- **THEN** exported native ChartML and fresh PPJ projection preserve the declared text/appearance and existing display flags

### Requirement: Source-bound label lifecycle
A recognized label SHALL support insertion, modification, deletion by omission and recreation under the existing analytics capability. No-op SHALL retain original source bytes. Label-only edits SHALL change only the target ChartPart and preserve other trendlines, error bars, data, axes and unrelated package parts.

#### Scenario: Remove and recreate
- **WHEN** a projected label is removed and later added again against the resulting source
- **THEN** each native owner and fresh projection reflect the requested state with no unrelated chart changes

### Requirement: Preserve unmodeled labels
Formula-backed text, nonempty manual layouts, source-linked formats, unsupported effects/extensions, duplicate nodes and malformed label content SHALL remain source-owned. These owners SHALL not gain analytics editing, and attempted replacement SHALL fail without emitting an edited file.

#### Scenario: Imported complex label
- **WHEN** an imported trendline label contains unsupported layout or text topology
- **THEN** unchanged compilation preserves its bytes and an analytics mutation is rejected
