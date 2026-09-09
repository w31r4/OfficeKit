## Purpose

Bind PPJ connector endpoints to explicit custom-shape connection-site indexes while retaining source identity and geometry dependencies.

## ADDED Requirements

### Requirement: Custom site endpoints preserve index identity

Connector from/to SHALL accept {element, connectionSite} with a zero-based index into a supported custom shape's connectionSites. It SHALL be mutually exclusive with coordinates and frame anchors. Authored endpoints SHALL resolve literal/reference site positions through target rotation/flips and ancestor coordinate transforms, and export native target/index bindings. Fresh projection SHALL preserve indexes. Missing/unsupported targets and out-of-range indexes MUST reject.

#### Scenario: Explicit custom site
- **WHEN** a connector references an existing custom site
- **THEN** export SHALL retain native target/index identity and computed coordinates, and fresh projection SHALL retain the PPJ endpoint

### Requirement: Source site edits follow semantic dependencies

Issued endpoint authority SHALL permit switching index/target, detaching or switching to frame anchors. Source no-op SHALL preserve bytes. Semantic target geometry/frame edits and supported frame native-leaf edits SHALL recompute dependent endpoint coordinates. Unsupported geometry native-leaf edits on a bound target MUST reject rather than retain stale coordinates. Unchanged endpoint, arrows, stroke and non-target package members SHALL be preserved. Preset/opaque site targets SHALL retain their existing source authority boundary.

#### Scenario: Change site or dependency
- **WHEN** separate original-source requests change a custom binding or its semantic geometry/frame dependency
- **THEN** native coordinates and fresh projection SHALL reflect the requested change without losing binding identity

### Requirement: Preview reports site binding limitations

Preview SHALL diagnose custom-site binding semantics that it cannot fully render.

#### Scenario: Unrendered binding
- **WHEN** a connector endpoint references a custom site
- **THEN** unsupported binding semantics SHALL NOT be reported as fully rendered
