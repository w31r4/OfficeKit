## ADDED Requirements

### Requirement: Typed inline formula
PPJ SHALL accept a formula run beside literal text runs and SHALL compile the
supported finite LaTeX subset to native editable PowerPoint Office Math.

#### Scenario: Formula is mixed with prose
- **WHEN** one paragraph contains a text run followed by a valid formula run
- **THEN** their order is retained and the formula is emitted as `a14:m`
  rather than a picture, SVG, font glyph or flattened text substitute

#### Scenario: Formula syntax is unsupported
- **WHEN** the source uses an unknown command, malformed group or exceeds a
  formula budget
- **THEN** validation rejects before writing a PPTX

### Requirement: Deterministic formula semantics
The compiler SHALL lower one finite AST to canonical OMML and SHALL prove that
its native output re-reads as the same AST.

#### Scenario: Authored equation is rebuilt and reimported
- **WHEN** an OfficeKit-authored PPTX retains its embedded program snapshot
- **THEN** PPJ recovery returns the exact original LaTeX and stable run IDs

#### Scenario: Third-party OMML is encountered
- **WHEN** a presentation contains Office Math outside the canonical profile or
  lacks a matching embedded PPJ
- **THEN** projection preserves it source-bound and does not fabricate LaTeX

### Requirement: Formula style boundary
Formula runs SHALL accept only the style properties that can be represented
consistently across the canonical Office Math graph.

#### Scenario: Formula uses unsupported text decoration
- **WHEN** a formula run declares a hyperlink, highlight, gradient, shadow,
  underline, strike, baseline, capitalization or arbitrary font override
- **THEN** semantic validation rejects instead of partially applying style
