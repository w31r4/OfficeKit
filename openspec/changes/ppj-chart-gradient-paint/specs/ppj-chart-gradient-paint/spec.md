# PPJ chart gradient-paint specification

## ADDED Requirements

### Requirement: Shared chart fill union

PPJ SHALL compile `none`, `solid`, and bounded direct-RGB `gradient` fills at
chart-area, plot-area, and series locations.

#### Scenario: Authored evidence chart

- **WHEN** a valid PPJ chart declares bounded chart and series fills
- **THEN** native output and re-import preserve fill kind, colors, stops,
  opacity and angle

### Requirement: Source-bound fill continuation

A canonical imported chart SHALL expose these locations through `setChartFill`;
changing only issued fill state SHALL patch its ChartPart, while unsupported or
stale fill graphs SHALL fail closed.

#### Scenario: Imported chart edit

- **WHEN** an Agent changes one chart surface and one series fill in a fresh PPJ
  projection
- **THEN** the second projection reports those fills without rebuilding the
  surrounding presentation

### Requirement: One discoverable fill model

Schema, compiler, capability registry, generated manual and focused chart
guidance SHALL describe the same fill union and native boundary.

#### Scenario: Maintainer check

- **WHEN** the Presentation Skill maintainer runs
- **THEN** generated PPJ documentation and capability metadata are synchronized
