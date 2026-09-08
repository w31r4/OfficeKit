# PPJ authored text glow

## ADDED Requirements

### Requirement: expose a bounded text glow field

The PPJ `textStyle` schema MUST accept a `glow` object using the existing
bounded glow profile. Authored paragraph default runs and authored text runs
MUST lower that object to one direct `a:glow` owner without changing text,
fill, typography, or hyperlink semantics.

#### Scenario: run glow compiles to the direct owner

- **WHEN** an authored rich-text run declares a glow with color `#D9A514`,
  radius `8`, and opacity `0.42`
- **THEN** its run properties contain one `a:glow` with the requested radius,
  color, and alpha
- **AND** embedded PPJ recovery retains the declarative glow fields

#### Scenario: default run glow keeps the paragraph owner

- **WHEN** a paragraph declares `defaultText.glow`
- **THEN** the paragraph's `a:defRPr` contains one direct `a:glow`
- **AND** a separate run glow or outer shadow is not merged into that owner

### Requirement: preserve the bounded effect boundary

Text glow lowering MUST remain independent from the existing bounded outer
shadow. When both are present, the canonical order MUST be `a:glow` followed
by `a:outerShdw`. Unsupported effect children MUST fail closed rather than be
dropped.

#### Scenario: invalid text glow fails closed

- **WHEN** a text glow declares a radius below `0`, above `1000`, or an
  opacity outside `0..1`
- **THEN** PPJ validation or native compilation fails

#### Scenario: imported text effects stay opaque

- **WHEN** an imported text run contains reflection, inner shadow, WordArt,
  or an arbitrary effect list
- **THEN** projection preserves the source-owned graph or refuses the edit
- **AND** it does not flatten the graph into PPJ `textStyle.glow`
