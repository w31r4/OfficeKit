# PPJ chart axis arrowheads

## ADDED Requirements

### Requirement: Ordinary native chart axes accept finite arrowheads

OfficeKit SHALL map bounded PPJ axis endpoint names to editable native
DrawingML arrowheads without introducing decorative overlay shapes.

#### Scenario: Author an increasing value axis

- **GIVEN** an ordinary native PPJ chart axis with an end arrow
- **WHEN** OfficeKit builds and reimports the presentation
- **THEN** the ChartPart contains the corresponding canonical `a:tailEnd`
- **AND** PPJ projection recovers the endpoint by stable semantic name

### Requirement: Canonical source axes continue locally

OfficeKit SHALL permit arrow changes only when the imported axis line graph is
canonical and the chart issues the existing axis capability.

#### Scenario: Change one source-bound endpoint

- **GIVEN** a canonical imported native chart with an editable axis line
- **WHEN** PPJ changes one `axisLineArrow` endpoint
- **THEN** only the existing ChartPart is patched
- **AND** a second projection reports the requested endpoint
- **AND** unrelated chart topology remains unchanged

### Requirement: Unsupported line graphs fail closed

OfficeKit SHALL reject arrowheads on series/grid lines, radar spokes and
generated vector axes, and SHALL preserve non-canonical native endpoints.

#### Scenario: Imported endpoint carries custom sizing

- **GIVEN** a third-party chart axis endpoint has explicit size metadata or an
  unknown child/effect
- **WHEN** OfficeKit projects the chart
- **THEN** it does not issue editable arrow semantics for that axis graph
- **AND** unrelated edits preserve the source content
