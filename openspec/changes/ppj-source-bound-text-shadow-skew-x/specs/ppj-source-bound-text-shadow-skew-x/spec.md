## ADDED Requirements

### Requirement: Direct rich-text run shadow horizontal skew

The presentation PPJ codec MUST expose a valid imported direct rich-text run outer shadow `outerShdw/@kx` as `run.style.shadow.skewX`, with native leaf kind `textShadowSkewX` and semantic degree precision of 1/60000.

#### Scenario: Project and edit an existing signed kx token

- **GIVEN** a source-bound presentation containing one text run with a bounded direct outer shadow, required blur/distance/direction, one RGB or theme color child, and an existing canonical signed `kx` token strictly between -5400000 and 5400000
- **WHEN** the presentation is projected to PPJ and a `textShadowSkewX` edit changes only that token
- **THEN** the projected style contains the signed degree value, the native leaf carries the degree value, the source-bound plan replaces only `outerShdw/@kx`, and the second projection reports the new value

#### Scenario: Preserve source-owned topology

- **GIVEN** the run is missing `kx`, has a non-canonical/invalid `kx`, has another transform (`sx`, `sy`, `ky`, or `rotWithShape`), lacks required geometry/color, or has sibling/unknown effect content
- **WHEN** the presentation is projected or a `textShadowSkewX` edit is attempted
- **THEN** the horizontal-skew leaf is absent or the operation fails closed, while unrelated source bytes remain preserved

#### Scenario: Reject stale or out-of-range edits

- **GIVEN** the source token no longer matches the projected precondition or an edit supplies a non-canonical/out-of-range signed native token
- **WHEN** the source-bound plan is compiled
- **THEN** compilation fails without rewriting the package
