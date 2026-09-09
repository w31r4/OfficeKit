## Purpose

Allow PPJ users to add, reorder and remove native trendlines after importing a chart while preserving its data, series identity and unrelated package content.

## ADDED Requirements

### Requirement: Writable ordered trendline list

On recognized native bar/column and line series, including categorical combo series, PPJ SHALL support insertion, reordering, replacement and removal through data.series[].trendlines. Omission or an empty array SHALL clear the list. The existing maximum count and per-trendline constraints SHALL remain enforced.

#### Scenario: Add and reorder trendlines
- **WHEN** the requested trendline array changes in length or order
- **THEN** native chart nodes and a fresh projection reflect the requested list in order

#### Scenario: Remove and recreate the list
- **WHEN** a projected series removes trendlines or sets an empty array, and later adds a new list
- **THEN** the native list is cleared and can subsequently be recreated through the same field

### Requirement: Localized native mutation

Trendline list edits SHALL change only the target ChartPart. The compiler SHALL preserve error bars, series identity, chart data and unrelated native content. All other ZIP entries SHALL retain their bytes.

#### Scenario: Edit one series in a combination chart
- **WHEN** one combo series adds or removes a trendline
- **THEN** other series, axes, data and error bars remain unchanged, and new trendlines precede error bars and data sources

### Requirement: Unsupported owners remain protected

Malformed or unsupported native trendline content SHALL NOT acquire editing capability through this change. A no-op round trip SHALL preserve it. The separate XLSX adapter SHALL retain its current trendline-count restriction.

#### Scenario: Unknown trendline label or extension
- **WHEN** an imported trendline contains a label, extension, duplicate scalar or other unsupported owner content
- **THEN** source-bound analytics editing is withheld and no-op compilation preserves the source ChartPart
