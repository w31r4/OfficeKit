## ADDED Requirements

### Requirement: Direct rich-text run shadow vertical scale

The presentation PPJ codec MUST expose a valid imported direct rich-text run outer shadow `outerShdw/@sy` as `run.style.shadow.scaleY`, with native leaf kind `textShadowScaleY` and semantic ratio precision of 0.00001.

#### Scenario: Project and edit an existing signed sy token

- **GIVEN** a source-bound presentation containing one text run with a bounded direct outer shadow, required blur/distance/direction, one RGB or theme color child, and an existing canonical signed `sy` token
- **WHEN** the presentation is projected to PPJ and a `textShadowScaleY` edit changes only that token
- **THEN** the projected style contains the signed ratio, the native leaf carries the ratio value, the source-bound plan replaces only `outerShdw/@sy`, and the second projection reports the new ratio

#### Scenario: Preserve source-owned topology

- **GIVEN** the run is missing `sy`, has a non-canonical/invalid `sy`, has both `sx` and `sy`, has skew or rotation, lacks required geometry/color, or has sibling/unknown effect content
- **WHEN** the presentation is projected or a `textShadowScaleY` edit is attempted
- **THEN** the vertical-scale leaf is absent or the operation fails closed, while unrelated source bytes remain preserved

#### Scenario: Reject stale or out-of-range edits

- **GIVEN** the source token no longer matches the projected precondition or an edit supplies a non-canonical/out-of-range signed native token
- **WHEN** the source-bound plan is compiled
- **THEN** compilation fails without rewriting the package
