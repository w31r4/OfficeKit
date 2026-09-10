## Purpose

Represent the two remaining native paragraph alignment modes in typed PPJ while preserving direct formatting, source edit authority and visible preview limitations.

## ADDED Requirements

### Requirement: Seven paragraph alignment modes

PPJ paragraph alignment SHALL accept left, center, right, justify, distributed, justifyLow and thaiDistributed. The last two SHALL serialize as native justLow and thaiDist, reproject under their PPJ names, and support the existing ordinary text/shape assignment, deletion, restoration and alignment-only wrapper removal lifecycle. Invalid spellings SHALL remain rejected and each source edit SHALL require exact field authority.

#### Scenario: Edit the remaining native alignment modes
- **WHEN** a paragraph is authored or source-edited with justifyLow or thaiDistributed
- **THEN** fresh native projection SHALL retain its exact mode, and removal/restoration SHALL preserve surrounding runs, paragraphs and non-target XML/ZIP

### Requirement: Shared paragraph owners retain the new modes

Supported table paragraphs and authored paragraph-default owners SHALL retain the new modes without degrading their existing topology merely because of alignment.

#### Scenario: Round trip shared paragraph properties
- **WHEN** a bounded table cell or master/list default is authored with either new mode
- **THEN** export and fresh projection SHALL retain that mode, and the table SHALL retain its bounded source editing capability

### Requirement: Explicit preview limitations

Schema, Help, registry and generated Agent guidance SHALL expose both names and their native mapping. Preview SHALL retain an explicit partial alignment/layout diagnostic for both modes until their glyph layout is implemented and verified.

#### Scenario: Preview a new alignment mode
- **WHEN** preview consumes either mode
- **THEN** its incomplete alignment painting SHALL remain machine-readable and SHALL NOT be reported as fully supported
