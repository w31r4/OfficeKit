## Purpose

Allow direct chart text reflections to retain their native transform attributes through PPJ authoring, projection and independent edits.

## ADDED Requirements

### Requirement: Reflection transform values preserve native meaning and presence

`chartTextStyle.reflection` SHALL accept optional `fadeAngle` in degrees [-360,360], normalized to [0,360); signed `scaleX/scaleY` ratios in [-21474.83648,21474.83647]; `skewX/skewY` in degrees strictly between -90 and 90; nine native `alignment` tokens tl/t/tr/l/ctr/r/bl/b/br; and boolean `rotateWithShape`. Values SHALL round ties to even to native 1/100000 ratio or 1/60000 degree precision, rejecting rounded skew at either excluded bound. Missing attributes, zero, negative scales and false SHALL remain distinct. Unknown keys, nulls and malformed native values SHALL not become editable values.

#### Scenario: Explicit values and defaults
- **WHEN** a chart reflection contains scaleX -1, scaleY 0, skewX -12.5, skewY 0, fadeAngle -90, alignment b and rotateWithShape false
- **THEN** native output and fresh PPJ projection retain every attribute, with fadeAngle 270
- **AND** replacing it with an empty object preserves the reflection with all transform attributes absent

#### Scenario: Boundaries
- **WHEN** authored or source-bound PPJ contains skew 90, rounded skew reaching 90, scale outside Int32 native range, invalid alignment or non-boolean rotation
- **THEN** compilation fails without an output artifact

### Requirement: Independent chart lifecycle and existing ownership boundaries

All existing chart text owners, whole-reflection style precedence and vector chart text defaults SHALL propagate these attributes. Editing or deleting reflection SHALL preserve other effects, literal text and non-target package entries. Fresh no-op compilation SHALL preserve original bytes. Ordinary imported reflection owners SHALL retain their existing full-span, no-transform guard. Unknown, duplicate or out-of-order effect graphs SHALL remain source-owned or fail closed.

#### Scenario: Source chart edit and vector override
- **WHEN** a projected line or combo chart changes reflection transforms and then removes and recreates the effect
- **THEN** each fresh projection reports the candidate attributes and only its chart part changes
- **AND** a vector title default carries chart reflection transforms while an explicit ordinary run reflection overrides the whole reflection object

#### Scenario: Unsupported ordinary owner
- **WHEN** ordinary imported text has a full-span reflection with an explicit transform attribute
- **THEN** it does not acquire the bounded ordinary reflection edit profile
