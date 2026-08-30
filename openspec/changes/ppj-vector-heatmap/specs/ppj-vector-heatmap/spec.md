## ADDED Requirements

### Requirement: PPJ SHALL declare a bounded heatmap matrix

PPJ SHALL accept `chartType: "heatmap"` only when categories and series names
form a unique non-empty rectangular matrix of no more than 32 by 32 cells and
every value is numeric or explicitly missing.

#### Scenario: valid matrix

- **WHEN** a heatmap has four string categories, three named series and each
  series has four numeric-or-null values
- **THEN** `ppj check` accepts the matrix shape before compilation

#### Scenario: ambiguous matrix

- **WHEN** categories or series names are duplicated, a value row has the wrong
  length, or the matrix exceeds the bounded dimensions
- **THEN** validation rejects the program without producing a PPTX

### Requirement: Heatmap colors SHALL be deterministic

The heatmap style SHALL use exactly two colors for a linear scale or exactly
three colors for a diverging scale. The compiler SHALL validate any explicit
domain and midpoint before interpolating colors.

#### Scenario: diverging scale

- **WHEN** a heatmap declares a negative-to-positive domain, three colors and a
  midpoint of zero
- **THEN** low, midpoint and high values compile to the corresponding endpoint
  colors and intermediate values interpolate deterministically

### Requirement: Heatmaps SHALL remain native editable vectors

The authored compiler SHALL lower a heatmap to one native DrawingML group with
editable cells, labels and colorbar. It SHALL NOT rasterize the matrix or claim
that it is a standard PowerPoint ChartPart.

#### Scenario: build and reopen

- **WHEN** a valid heatmap PPJ is built and reopened with its embedded program
- **THEN** OfficeKit restores the exact heatmap PPJ and the PPTX contains the
  expected native group children with no chart or image part for the heatmap

#### Scenario: embedded program removed

- **WHEN** the same PPTX is imported after its OfficeKit program parts are
  removed
- **THEN** OfficeKit projects the native group and does not infer heatmap
  semantics from arbitrary DrawingML shapes

### Requirement: Unsupported heatmap behavior SHALL fail closed

The bounded heatmap SHALL reject secondary axes, per-series chart overrides,
trendlines, error bars, markers, chart-build animation and other ordinary
ChartPart-only settings.

#### Scenario: chart build requested

- **WHEN** an animation requests series/category chart build for a vector
  heatmap
- **THEN** validation rejects the request and recommends whole-group reveal
