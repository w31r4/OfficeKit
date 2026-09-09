## Purpose

Expose existing error-bar formula bindings in PPJ and keep edited numeric error amounts consistent between the chart cache and its embedded worksheet.

## ADDED Requirements

### Requirement: Source-bound formula identity
Complete imported custom error data using local single-row or single-column ranges SHALL project its `formula`, `values` and optional `formatCode`. Formula identity SHALL remain fixed for this operation; creating, deleting or retargeting a formula SHALL be rejected, as shall source-free formula authoring.

#### Scenario: Inspect and preserve a formula
- **WHEN** a supported source error side contains a complete numeric cache and local range formula
- **THEN** PPJ exposes both without converting the reference to literals, and a no-op preserves the original bytes

### Requirement: Atomic error amount update
Changing error values within an existing formula binding SHALL update matching native cache points and worksheet cells together. This applies to column/bar/line and categorical combo series. Zero values SHALL remain valid. The export SHALL preserve all other outer and inner package parts and relationships.

#### Scenario: Update both sides
- **WHEN** plus and minus error amounts change in uniquely owned, matching numeric worksheet cells
- **THEN** the chart and worksheet contain the new numbers, fresh PPJ projection recovers them, and changed-part evidence identifies the ChartPart and embedded workbook

### Requirement: Fail closed on unresolved workbook ownership
The operation SHALL reject missing or shared workbook bindings, stale cache/cell values, overlapping chart consumers, formula cells and unsupported dependency graphs. A failed compilation SHALL produce no edited file. Style-only changes SHALL not rewrite workbook data.

#### Scenario: Unproven cell closure
- **WHEN** the referenced cell is shared by another chart channel, contains a formula, disagrees with the cache, or the workbook is shared
- **THEN** a value edit fails while an unchanged source remains preserved
