## ADDED Requirements

### Requirement: Formal text precedence sources

The PPJ grammar MUST accept the source names `theme`, `master`, `layout`,
`styleRef`, `element`, `paragraph`, `run`, and `default` in addition to the
legacy `inline` source. A rule MUST continue to use unique sources and the
first declared source with a value MUST win.

#### Scenario: A declared run value wins

- **GIVEN** a `text.size` rule declares `run`, `paragraph`, `element`,
  `styleRef`, `layout`, `master`, and `default` in that order
- **AND** more than one owner supplies a size
- **WHEN** the authored compiler resolves a run
- **THEN** it MUST use the direct run value
- **AND** the projected native run MUST contain that effective size

#### Scenario: Missing specific owners fall through

- **GIVEN** a `text.size` rule declares the formal source order
- **AND** the run and paragraph do not supply a size
- **WHEN** the compiler resolves the run
- **THEN** it MUST select the first available element, named-style, layout,
  master, theme, or default value in the declared order

### Requirement: Concrete PPJ hierarchy owners

Text and placeholder elements MAY declare a direct `textStyle`; design master
and layout definitions MAY declare a direct `style` text-style fallback. The
compiler MUST keep the existing text-container `style` semantics separate
from the direct element text-style owner.

#### Scenario: Element text style is distinct from body layout

- **GIVEN** a text element declares both `style` body properties and
  `textStyle.size`
- **WHEN** the element is authored
- **THEN** `textStyle.size` MUST participate in `text.size` precedence
- **AND** the body properties MUST still control only the text container

### Requirement: Bounded authored lowering

For `text.size`, `text.bold`, `text.italic`, `text.font`, and
`text.fontFamily`, the authored compiler MUST resolve the formal sources per
run and lower only the effective value to the existing native run owner. It
MUST NOT claim source-bound master/layout editing or create a new native style
graph as a side effect.

#### Scenario: Effective value survives projection

- **GIVEN** a valid PPJ using formal text precedence
- **WHEN** it is compiled, the embedded PPJ is removed, and the PPTX is
  projected back to PPJ
- **THEN** the effective run scalar values MUST be recoverable
- **AND** the projection MUST not claim the discarded declaration graph is
  source-bound editable

### Requirement: Review uses the same owners

The read-only grammar evaluator MUST recognize the formal source names and
report the selected source for each rich-text run, including the page's
layout/master fallback owners.

#### Scenario: Review and compiler agree on the selected source

- **GIVEN** a formal `text.size` rule and a run with values at multiple layers
- **WHEN** the PPJ is reviewed
- **THEN** the resolution entry MUST identify the same winning source as the
  authored compiler
