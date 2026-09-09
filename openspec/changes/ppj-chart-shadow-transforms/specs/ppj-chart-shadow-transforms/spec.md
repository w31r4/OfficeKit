## Purpose

Let PPJ preserve direct chart text outer-shadow scaling and skew without discarding independent effects or source package content.

## ADDED Requirements

### Requirement: Optional chart shadow transforms

`chartTextStyle.shadow` SHALL accept optional signed ratios `scaleX/scaleY` in [-21474.83648,21474.83647] and degree angles `skewX/skewY` strictly between -90 and 90. Ratios SHALL retain 1/100000 precision and angles 1/60000 degree precision, with ties-to-even rounding. Rounded skew at either excluded bound SHALL fail compilation. Zero, negative values and omission SHALL remain distinct. Existing color, geometry, alignment, rotation and opacity behavior SHALL remain intact.

#### Scenario: Signed transforms and clearing
- **WHEN** a chart shadow with color #112233 has scaleX -1, scaleY 0, skewX -12.5 and skewY 0
- **THEN** authored output and fresh projection retain all values
- **AND** removing those properties preserves the color and shadow while removing only their direct native attributes

#### Scenario: Invalid values
- **WHEN** authored or source-bound input has out-of-range scale, skew equal to or rounding to ±90, null or unknown transform keys
- **THEN** compilation fails without an output artifact

### Requirement: Source lifecycle and shared chart styles

All existing chart text owners, whole-shadow style precedence and vector defaults SHALL carry these attributes. Source edits, shadow deletion and recreation SHALL preserve sibling effects, literal text and non-target ZIP entries. Fresh no-op SHALL return identical bytes. Unknown effect topology SHALL remain source-owned or fail closed; ordinary imported shadow owners SHALL retain their no-transform guard.

#### Scenario: Line and combo lifecycle
- **WHEN** a projected line or combo chart changes and clears transforms, deletes and recreates the shadow
- **THEN** each fresh projection recovers the candidate and only the owning chart part changes
- **AND** other effects and original literal text remain unchanged

#### Scenario: Style overrides and ordinary guard
- **WHEN** a vector title default contains a transformed shadow and a run declares an ordinary shadow override
- **THEN** the default carries all transforms and the explicit run replaces that entire shadow
- **AND** an ordinary imported shadow with even a default-valued transform remains outside its existing scalar edit profile
