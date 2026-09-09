## Purpose

Complete the PPJ text-body rotation lifecycle while preserving explicit zero, signed angles and native absence across source edits.

## ADDED Requirements

### Requirement: Text-body rotation can be removed and restored

Existing editable text, shape, master/layout placeholder and table-cell text styles SHALL preserve explicit rotation degrees and omission. Removing a previously projected rotation SHALL remove the direct bodyPr rot attribute. Explicit zero SHALL remain a native value. A whole style containing only rotation and/or upright SHALL be removable. Source authority and other whole-style removal guards MUST remain enforced.

#### Scenario: Remove then restore rotation
- **WHEN** an authorized source request removes rotation and a fresh-source request adds zero or a signed angle
- **THEN** native attributes and fresh PPJ SHALL preserve the requested presence and value, including compact table text restoration

### Requirement: Rotation edits preserve unrelated state

Source no-op SHALL preserve bytes. Rotation edits SHALL preserve element frame rotation, other text-body values, text topology and non-target ZIP bytes. Preview SHALL continue reporting unimplemented body rotation layout semantics.

#### Scenario: Change body rotation only
- **WHEN** a source edit changes only text-body rotation
- **THEN** the native body attribute SHALL change without modifying unrelated modeled state
