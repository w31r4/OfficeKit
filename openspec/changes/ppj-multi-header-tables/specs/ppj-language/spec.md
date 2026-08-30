## ADDED Requirements

### Requirement: Authored PPJ tables support bounded multi-row headers

The authored PPJ compiler SHALL accept a `table.style.headerRows` value from
zero through the table's physical row count. It SHALL apply cell-local styling
first, header styling second, and table defaults last. A non-zero header count
SHALL set the native first-row flag while direct cell formatting SHALL preserve
the intended appearance of every declared header row.

#### Scenario: Two-level native header

- **WHEN** a source-free PPJ table declares two header rows plus
  `headerCellFill` and `headerTextStyle`
- **THEN** both header rows compile to editable native table cells with those
  fallbacks
- **AND** an explicit cell fill or text style still wins
- **AND** the embedded PPJ recovers `headerRows: 2` exactly

#### Scenario: Invalid header declaration

- **WHEN** `headerRows` exceeds the physical row count
- **OR** header-only styles are present while `headerRows` is zero
- **THEN** compilation fails before writing output with a path-specific error
