# PPJ category combo charts

## ADDED Requirements

### Requirement: PPJ authors bounded categorical combo plots

PPJ SHALL author one native editable combo ChartPart from two or three distinct
column, line and area plot families over one aligned category domain.

#### Scenario: Area, column and line share a category domain

- **WHEN** a valid PPJ combo assigns each complete plot family to a primary or
  secondary axis pair
- **THEN** the compiler SHALL emit native plot elements with stable series
  order, shared categories and deterministic axis references
- **AND** import SHALL recover the same series types, axis groups and values.

#### Scenario: Numeric or inconsistent combo is requested

- **WHEN** a combo requests scatter, bubble, horizontal bar, only one plot
  family or splits one family across axis pairs
- **THEN** validation SHALL reject it before writing output
- **AND** an imported unsupported graph SHALL remain source-preserved.

