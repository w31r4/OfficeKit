## ADDED Requirements

### Requirement: Design profiles are source-bound and deterministic
The profile generator SHALL accept one or more PPTX files and emit a versioned
profile whose source filename, byte length, SHA-256, and structural evidence
are stable for the same input bytes.

#### Scenario: Repeat profiling of one deck
- **WHEN** the same PPTX bytes are profiled twice
- **THEN** the complete JSON profile is byte-for-byte identical and contains no
  absolute input path

### Requirement: Profiles expose evidence-backed design language
The profile SHALL report canvas, source XML palette/typeface/size evidence,
slide density, layout families, slide archetypes, recurring geometry
candidates, and native opaque-object summaries without copying source package
bytes into the profile.

#### Scenario: Profile a deck with opaque native objects
- **WHEN** an imported deck contains SmartArt, OLE, WPS, or another native
  object
- **THEN** the profile identifies the object kind and stable inspected ID while
  leaving it opaque and preserving the source hash

### Requirement: Design evidence does not grant mutation authority
The profile SHALL NOT authorize an edit, clone, or raw OOXML mutation. Any
subsequent operation MUST use the existing source-bound typed or
capability-issued preconditions.

#### Scenario: Agent selects a recurring component
- **WHEN** a recurring geometry candidate is selected from a profile
- **THEN** the operation is rejected unless the source revision and ownership
  proof independently pass the existing codec checks
