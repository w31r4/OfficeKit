## ADDED Requirements

### Requirement: Imported slide overlay discovery
An ordinary imported page SHALL advertise bounded typed overlay authority in
its page nativeRef.

#### Scenario: Agent inspects a continuable page
- **WHEN** a PPTX slide is projected as an ordinary source-bound PPJ page
- **THEN** the page advertises `appendElement` for `elements`

### Requirement: Declarative typed overlay
A source-bound PPJ SHALL express a new overlay as typed elements appended to
the end of the page's ordered element array.

#### Scenario: Agent adds a source-preserving label
- **WHEN** an Agent retains every source element in order and appends a fresh
  textbox authorized by the page capability
- **THEN** build adds a native editable textbox above the preserved source
  prefix through the existing overlay writer

#### Scenario: Agent adds a sourced image
- **WHEN** an Agent declares and supplies a hash-bound image asset and appends
  a fresh image element
- **THEN** build embeds the image and adds only the bounded slide relationship
  and media part required by that overlay

### Requirement: Overlay profile is finite
The compiler SHALL admit only typed text, rect, roundRect, ellipse, and image
elements in the source-bound overlay suffix.

#### Scenario: Agent appends a chart or arbitrary shape
- **WHEN** the new suffix contains a chart, table, group, connector, media,
  placeholder, diagram, OLE, opaque object, custom geometry, or another preset
  shape
- **THEN** check or build rejects the program before changing the source PPTX

### Requirement: Source prefix and mutation isolation
The compiler SHALL preserve the exact source element prefix and SHALL reject
an overlay combined with another mutation of that source slide.

#### Scenario: Agent interleaves an overlay
- **WHEN** a new element appears before or between source-bound elements
- **THEN** build rejects the z-order request instead of moving unknown native
  content

#### Scenario: Agent edits and appends together
- **WHEN** the same build changes a source element, deletes or reorders it, or
  changes slide metadata while appending an overlay
- **THEN** build rejects the mixed transaction and asks for a build/reimport
  boundary

### Requirement: Reimport yields ordinary source state
An appended overlay SHALL become an ordinary projected source-bound element
after build and reimport.

#### Scenario: Agent continues after an overlay build
- **WHEN** the output PPTX is projected again
- **THEN** the new element has a stable typed PPJ node and nativeRef while all
  unrelated source content remains available

### Requirement: Agent discoverability
The generated PPJ reference SHALL explain capability discovery, suffix-only
placement, supported element kinds, asset requirements, and the mandatory
build/reimport boundary.

#### Scenario: Fresh Agent continues an imported page
- **WHEN** an Agent reads the PPJ import guide
- **THEN** it can add a bounded overlay without MJS, raw OOXML, part paths, or
  a procedural mutation list
