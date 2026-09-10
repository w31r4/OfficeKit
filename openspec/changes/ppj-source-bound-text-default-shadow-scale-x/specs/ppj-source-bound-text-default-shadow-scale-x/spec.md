## Purpose

This capability lets an imported paragraph default-text outer shadow change its
horizontal scale in PPJ while preserving the source effect graph and package.

## ADDED Requirements

### Requirement: project paragraph default-shadow horizontal scale

The PPJ projector MUST expose `text.paragraphs[].style.defaultText.shadow.scaleX`
through the existing `setTextParagraphStyle` capability and a native leaf when
an editable text shape contains one strict direct default-run outer shadow with
an existing bounded `sx` attribute.

#### Scenario: default shadow scale receives a native leaf

- **WHEN** an imported text shape has a direct paragraph default `a:outerShdw`
  with a valid `sx` value
- **THEN** the element reference contains the default-text style capability
  and a leaf whose native value is the source `sx` token and whose PPJ value is
  the corresponding scale ratio

### Requirement: edit only the default-shadow scale token

The source-bound compiler MUST accept a changed finite scale ratio in the
DrawingML signed Int32 range, round it to native 1/100000 precision, and change
only the owning slide's direct `a:outerShdw/@sx` attribute.

#### Scenario: scale-only edit round-trips

- **WHEN** a projected paragraph default shadow changes `scaleX` while its
  color, opacity, other geometry, effects, paragraph/run topology, and text
  remain unchanged
- **THEN** compilation succeeds, only the owning SlidePart changes, package
  XML remains valid, and a fresh projection reports the new scale and all
  preserved sibling state

### Requirement: reject unsafe default-shadow scale edits

The compiler MUST fail closed for missing or malformed `sx`, values outside the
bounded signed scale range, stale leaf authority, unsupported effect graphs,
and edits that also change shadow identity, text topology, or unrelated source
  ownership.

#### Scenario: unsupported default-shadow transform is preserved

- **WHEN** the source shadow has duplicate effect lists, unknown descendants,
  or an invalid `sx` token and a request attempts to replace its scale
- **THEN** source-bound compilation fails without mutating the source package
