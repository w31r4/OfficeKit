## Purpose

Let PPJ express native trendline label manual layout with explicit field presence and a source-preserving edit lifecycle.

## ADDED Requirements

### Requirement: Typed manual layout

PPJ SHALL expose optional `label.layout.manual` with optional `target` (inner/outer), `xMode`, `yMode`, `widthMode`, `heightMode` (edge/factor), and finite numeric `x`, `y`, `width`, `height`. Values SHALL remain native chart fractions without clamping or slide-unit conversion. Omitted layout, empty layout and empty manual layout SHALL remain distinguishable.

#### Scenario: Native layout and presence round trip
- **WHEN** an ordinary or categorical combo chart is authored with a manual layout containing modes, target, zero, negative offsets and a value above one
- **THEN** native chart layout and fresh projection SHALL preserve those values and optional presence
- **AND** empty layout and empty manual layout SHALL round trip distinctly

### Requirement: Source-bound layout lifecycle

Supported trendline labels SHALL allow layout addition, modification, reset to either empty container, deletion and recreation through the existing analytics edit operation. A no-op SHALL preserve all source bytes; a layout-only edit SHALL preserve all non-target ZIP entries and chart content outside the label layout. Present native mode/target leaves without a value SHALL project their standard factor/outer defaults.

#### Scenario: Layout editing leaves label and chart state intact
- **WHEN** a freshly projected chart's label layout is edited, removed and recreated against its original bytes
- **THEN** only the target ChartPart SHALL change, existing label content/style SHALL remain intact, and fresh projection SHALL equal the requested layout

### Requirement: Unsupported layout content remains safe

Unknown layout content and malformed native layouts SHALL remain source-owned without analytics editing. Invalid PPJ types, modes or numeric values SHALL fail without producing an output file.

#### Scenario: An extension cannot be erased by an analytics edit
- **WHEN** a source manual layout contains an extension, unknown attribute, duplicate field, invalid ordering or malformed value
- **THEN** no-op SHALL preserve source bytes and attempted analytics replacement SHALL fail without output
