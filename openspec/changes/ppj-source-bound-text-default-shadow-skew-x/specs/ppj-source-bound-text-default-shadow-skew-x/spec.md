## Purpose

This capability lets an imported paragraph default-text outer shadow change
its horizontal skew in PPJ while preserving the source effect graph and package.

## ADDED Requirements

### Requirement: project paragraph default-shadow horizontal skew

The PPJ projector MUST expose
`text.paragraphs[].style.defaultText.shadow.skewX` through the existing
`setTextParagraphStyle` capability and a native leaf when an editable text
shape contains one strict direct default-run outer shadow with an existing
bounded `kx` attribute.

#### Scenario: default shadow skew receives a native leaf

- **WHEN** an imported text shape has a direct paragraph default `a:outerShdw`
  with a valid `kx` value
- **THEN** the element reference contains the default-text style capability and
  a leaf whose native value is the source `kx` token and whose PPJ value is the
  corresponding degree value at 1/60000 precision

### Requirement: edit only the default-shadow horizontal skew token

The source-bound compiler MUST accept a changed finite skew value whose native
1/60000-degree integer remains strictly between -5400000 and 5400000, and
change only the owning slide's direct `a:outerShdw/@kx` attribute.

#### Scenario: skew-only edit round-trips

- **WHEN** a projected paragraph default shadow changes `skewX` while its
  color, opacity, other geometry, effects, paragraph/run topology, and text
  remain unchanged
- **THEN** compilation succeeds, only the owning SlidePart changes, package XML
  remains valid, and a fresh projection reports the new skew and all preserved
  sibling state

### Requirement: reject unsafe default-shadow skew edits

The compiler MUST fail closed for missing or malformed `kx`, values whose
rounded native integer is at either excluded bound, stale leaf authority,
unsupported effect graphs, and edits that also change shadow identity, text
topology, or unrelated source ownership.

#### Scenario: unsupported default-shadow skew is preserved

- **WHEN** the source shadow has duplicate effect lists, unknown descendants,
  an invalid `kx` token, or a bound `kx` token and a request attempts to
  replace its skew
- **THEN** source-bound compilation fails without mutating the source package
