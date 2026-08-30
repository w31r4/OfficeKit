## ADDED Requirements

### Requirement: PPJ declares waterfall point semantics
The PPJ language SHALL represent one waterfall series with a `pointRoles` array parallel to the category and value arrays, where each role is `delta` or `total`.

#### Scenario: Valid cumulative bridge
- **WHEN** one waterfall series has equal category, value, and point-role counts and its running total remains non-negative
- **THEN** PPJ validation accepts the program and retains every ordered value and role

#### Scenario: Ambiguous or inconsistent bridge
- **WHEN** a waterfall has multiple series, missing values, mismatched point-role count, a cumulative value below zero, or a later total inconsistent with the computed running value
- **THEN** validation fails before compilation with a path-specific diagnostic

### Requirement: PPJ compiles waterfall charts as native editable charts
The authored compiler SHALL lower a valid semantic waterfall into one standard native stacked-column chart with an invisible offset series and distinct increase, decrease, and total series.

#### Scenario: Authored waterfall build
- **WHEN** a valid waterfall declares explicit role fills and optional role strokes
- **THEN** the output ChartPart contains four deterministic stacked-column series, exactly one visible role at each category, and preserves the declared frame, axes, surfaces, title, and accessibility state

#### Scenario: Unsupported waterfall option
- **WHEN** a waterfall requests an image role fill, legend, explicit stacking, automatic data labels, trendline, error bar, marker, secondary axis, or another unsupported chart behavior
- **THEN** the compiler rejects the request before package output rather than dropping or approximating the option

### Requirement: Authored waterfall intent recovers exactly
OfficeKit SHALL recover the original semantic waterfall PPJ from an authored PPTX whose embedded program and map remain valid.

#### Scenario: Exact authored recovery
- **WHEN** an OfficeKit-authored waterfall PPTX is imported without invalidating its embedded program
- **THEN** the recovered PPJ still has `chartType: "waterfall"`, the original single series, ordered point roles, and role styles

#### Scenario: Ordinary third-party import
- **WHEN** a PPTX lacks a valid embedded PPJ and contains a standard stacked bridge or a ChartEx waterfall
- **THEN** OfficeKit describes only the native graph it can prove and keeps unsupported waterfall topology source-owned without inferring PPJ waterfall intent
