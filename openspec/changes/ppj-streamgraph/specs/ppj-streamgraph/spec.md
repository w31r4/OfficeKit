## ADDED Requirements

### Requirement: Authored PPJ streamgraph
PPJ SHALL express a bounded centered streamgraph through area-chart stream
stacking and compile it to editable native DrawingML.

#### Scenario: Valid streamgraph is authored
- **WHEN** an area chart declares `style.stacking: "stream"` with aligned
  ordered non-negative series
- **THEN** the compiler emits one stable editable band per series plus bounded
  category and series labels

#### Scenario: Streamgraph is rebuilt deterministically
- **WHEN** the same PPJ is compiled repeatedly
- **THEN** the PPTX program and generated native path graph are deterministic

#### Scenario: Authored streamgraph is reimported
- **WHEN** an OfficeKit-authored streamgraph PPTX retains its embedded program
- **THEN** PPJ import restores the exact area chart, stable IDs, values and
  `stacking: "stream"`

### Requirement: Streamgraph boundaries fail closed
PPJ SHALL reject stream semantics that cannot be represented truthfully by the
bounded editable vector compiler.

#### Scenario: Stream is used on another chart family
- **WHEN** a non-area chart declares `style.stacking: "stream"`
- **THEN** validation rejects before output is written

#### Scenario: Stream data is invalid
- **WHEN** a streamgraph has a negative value, an empty category total, or an
  unsupported series option
- **THEN** compilation rejects instead of silently changing the data meaning

#### Scenario: Third-party group resembles a streamgraph
- **WHEN** an imported PPTX contains arbitrary layered custom paths
- **THEN** OfficeKit preserves or projects those shapes without inferring PPJ
  stream semantics or issuing a stream edit capability

