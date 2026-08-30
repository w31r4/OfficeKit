## ADDED Requirements

### Requirement: Complete base transition vocabulary
PPJ SHALL express every canonical base slide-transition effect already owned by
the Presentation codec, including its applicable direction, orientation,
speed, through-black, spokes, duration, and advance state.

#### Scenario: Timed split transition
- **WHEN** a page declares a horizontal inward split, fast speed, 750 ms
  duration, click advance disabled, and timed advance after 1,250 ms
- **THEN** build emits one native editable transition and reimport recovers
  every declared value

### Requirement: Effect-specific validation
The PPJ validator SHALL reject transition properties that cannot be represented
by the selected effect instead of silently discarding them.

#### Scenario: Direction on a circle transition
- **WHEN** a circle transition also declares a cardinal direction
- **THEN** `check` and `build` reject the program with a diagnostic naming the
  incompatible field and effect

### Requirement: Conservative imported projection
The PPTX projector SHALL recover every supported canonical base-transition
field and SHALL advertise mutation only when the source codec proves the slide
transition editable or addable.

#### Scenario: Imported wheel transition
- **WHEN** a third-party page contains a supported six-spoke wheel transition
- **THEN** PPJ contains `type: wheel`, `spokes: 6`, its timing and advance state,
  and a source-bound transition capability

### Requirement: Source-bound local transition edit
A source-bound PPJ transition change SHALL use the existing transition
capability and SHALL mutate only the target SlidePart.

#### Scenario: Change an imported page to zoom out
- **WHEN** an Agent changes one capable source page from a base transition to
  `zoom` with `direction: out`
- **THEN** build preserves all non-target OPC parts and reimport recovers the
  new transition without rebuilding the page

### Requirement: Morph remains separate
PPJ Morph SHALL continue to require an adjacent source page and compatible
object pairs, and SHALL reject base-only transition fields.

#### Scenario: Morph with wheel spokes
- **WHEN** a Morph transition declares `spokes`
- **THEN** validation rejects the program before native compilation

### Requirement: Agent discoverability
The generated PPJ reference and Motion guidance SHALL document the complete
language while distinguishing communication-safe defaults from explicit
high-noise effects.

#### Scenario: Agent searches for an automatic timed transition
- **WHEN** an Agent reads PPJ transition guidance
- **THEN** it can find the fields for click and timed advance without reading
  protobuf or C# source
