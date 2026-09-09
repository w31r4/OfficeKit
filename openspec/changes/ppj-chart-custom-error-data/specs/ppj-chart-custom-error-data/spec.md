## Purpose

Express per-point upper and lower error amounts directly in PPJ so local custom chart uncertainty can be created and edited without losing native source ownership.

## ADDED Requirements

### Requirement: Literal custom error data
PPJ SHALL accept `valueType: "custom"` with plus/minus objects containing a required `values` array and optional `formatCode`. Each array SHALL contain one finite nonnegative JavaScript-safe number for every series point, including zero. `both` (the default) SHALL require both sides; `plus` and `minus` SHALL require only their selected side and reject the excluded side. Custom mode SHALL reject scalar `value`; other modes SHALL reject plus/minus data. Format codes SHALL contain 1–255 characters without controls and accept a string grammar token.

#### Scenario: Asymmetric point errors
- **WHEN** a bar/column, line or categorical combo series specifies different plus and minus arrays
- **THEN** native export and fresh projection retain each amount, side, optional format and existing cap/stroke fields

#### Scenario: Invalid data
- **WHEN** arrays have incorrect lengths, missing sides, excluded sides, negative or null values, or invalid format codes
- **THEN** compilation reports an error rather than truncating, filling or silently dropping data

### Requirement: Complete local object lifecycle
Imported literal custom error data SHALL support modifying values/formats, switching both/plus/minus, conversion to and from scalar modes, removal and recreation through the existing analytics capability. Each edit SHALL preserve other chart owners and unrelated ZIP entries.

#### Scenario: Source-bound lifecycle
- **WHEN** a projected custom object is edited, changed to one-sided data, replaced with a scalar, removed and recreated
- **THEN** each export and fresh projection reflects the requested object, and only the selected error-bar owner changes

### Requirement: Preserve formula and unsupported owners
Formula-backed custom error data SHALL remain source-owned and SHALL NOT be projected as local literal arrays. Attempted insertion over an unprojected owner SHALL fail closed. Malformed native caches SHALL not gain analytics editing capability. Local SVG preview SHALL retain an explicit partial error-bars-not-rendered diagnostic.

#### Scenario: Formula or malformed cache
- **WHEN** a source contains a custom formula or an unsupported cache topology
- **THEN** a no-op preserves it and the literal-data editing path cannot overwrite its source semantics
