# Specification: source-bound table last-row emphasis

## ADDED Requirements

### Requirement: expose an existing direct table last-row flag

The PPJ projection MUST expose `tableLastRow` only for a recognized rectangular
table whose direct `a:tblPr` contains exactly one canonical `lastRow` attribute
with value `0` or `1`. The projected table style MUST preserve the same value
at `table.style.lastRow`.

#### Scenario: project a direct last-row flag

- GIVEN an imported rectangular table with `a:tblPr lastRow="1"`
- WHEN the PPTX is projected to PPJ
- THEN the table style contains `lastRow: true`
- AND its native reference contains one `tableLastRow` leaf with value `1`

### Requirement: edit only the selected table-property token

The source-bound compiler MUST accept a changed canonical `0`/`1` value for a
capability-issued `tableLastRow` leaf and splice only the existing `lastRow`
attribute value in the owning table property element.

#### Scenario: edit last-row emphasis

- GIVEN a capability-issued `tableLastRow` leaf with expected value `1`
- WHEN its value is changed to `0`
- THEN only the owning slide part changes
- AND the table grid, cell content, relationships, and other table flags remain unchanged

### Requirement: fail closed for missing or irregular source

The codec MUST NOT issue or apply `tableLastRow` when the direct attribute is
absent, duplicated, malformed, non-canonical, or the table topology is outside
the recognized rectangular profile.

#### Scenario: absent last-row flag stays source-owned

- GIVEN an imported rectangular table with no direct `a:tblPr/@lastRow`
- WHEN the PPTX is projected to PPJ
- THEN no `tableLastRow` leaf is emitted
