## Purpose

This capability gives imported rich-text runs a bounded PPJ address for the horizontal scale of their direct outer shadow while preserving the original Open XML owner and surrounding effect graph.

## ADDED Requirements

### Requirement: Project a strict direct-run outer-shadow scaleX leaf

The system SHALL expose `run.style.shadow.scaleX` and the opaque native leaf `textShadowScaleX` when a text run owns exactly one direct `a:outerShdw` with an existing canonical signed Int32 `sx` token, valid bounded shadow geometry, one direct RGB or theme color, and no `sy`, `kx`, `ky`, or `rotWithShape` transform. The PPJ value SHALL be the signed native token divided by 100000.

#### Scenario: Imported scaleX is projected

- **WHEN** an imported text run contains a strict direct outer shadow with `sx="-50000"`, bounded geometry, and a direct color
- **THEN** the projection SHALL contain `run.style.shadow.scaleX` equal to `-0.5` and one `textShadowScaleX` leaf whose semantic value is `-0.5` (the bound native token is `-50000`)

#### Scenario: Unsupported shadow topology stays opaque

- **WHEN** the direct shadow is missing `sx`, has a non-canonical or out-of-range `sx`, has another shadow transform, has a sibling effect, or has an unknown child
- **THEN** the projection SHALL omit `textShadowScaleX` and SHALL NOT claim that graph is editable through this capability

### Requirement: Edit only the proven scale token

The system SHALL accept a source-bound `textShadowScaleX` edit only when the expected and requested values are changed canonical signed Int32 tokens, the source proof still identifies the same run and shadow, and the target owner remains strict. The edit SHALL replace only `a:outerShdw/@sx` in the owning SlidePart and SHALL preserve the other shadow fields, effect order, run/paragraph topology, and non-target package parts.

#### Scenario: Source-bound scaleX edit round-trips

- **WHEN** a projected leaf with semantic value `-0.5` is edited to `1.25`
- **THEN** compilation SHALL change only the owning SlidePart's `outerShdw/@sx` from `-50000` to `125000`, report that changed part, and a second projection SHALL report `scaleX` and the leaf semantic value equal to `1.25`

#### Scenario: Stale or invalid scaleX edit fails closed

- **WHEN** the raw token no longer matches the expected value or either value is not a canonical signed Int32 token
- **THEN** compilation SHALL fail closed without rewriting the package
