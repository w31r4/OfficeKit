## Purpose

Expose the direct rich-text reflection alignment value in PPJ with a bounded native leaf and a source-preserving edit path for canonical DrawingML owners.

## ADDED Requirements

### Requirement: Direct run reflection alignment is typed

The PPJ v1 schema SHALL accept `run.style.reflection.alignment` using only `tl`, `t`, `tr`, `l`, `ctr`, `r`, `bl`, `b`, or `br`, and an imported direct rich-text run SHALL expose that value as `textReflectionAlignment` only for a full-span direct reflection with one canonical `algn` and no fade, scale, skew, or rotate-with-shape attributes.

#### Scenario: Canonical alignment projects
- **WHEN** a run-owned `a:reflection` has `algn="b"`, full-span endpoints, and no other reflection transform
- **THEN** the run style contains `reflection.alignment = "b"` and one native leaf of kind `textReflectionAlignment` with value `b`

#### Scenario: Unsupported alignment topology stays opaque
- **WHEN** the run reflection also has a scale, skew, fade direction, rotate-with-shape flag, variable endpoint, duplicate attribute, or unsupported effect sibling
- **THEN** the alignment field and native leaf are not issued for that source owner

### Requirement: Alignment edit is source-bound

A source-bound PPJ edit SHALL accept a changed canonical alignment token, prove the same run-owned reflection and leaf value, replace only `a:reflection/@algn` in the owning SlidePart, preserve unrelated package bytes, and make the new token visible on a subsequent projection.

#### Scenario: Canonical alignment token is replaced
- **WHEN** a `textReflectionAlignment` leaf changes from `b` to `tr`
- **THEN** only the owning SlidePart changes, the Open XML reflection has `algn="tr"`, and the next PPJ projection reports `reflection.alignment = "tr"`

#### Scenario: Invalid or stale edit is rejected
- **WHEN** the requested token is outside the nine-value enum, is unchanged, or the source attribute no longer matches the expected leaf value
- **THEN** compilation rejects the edit without rewriting unrelated source content
