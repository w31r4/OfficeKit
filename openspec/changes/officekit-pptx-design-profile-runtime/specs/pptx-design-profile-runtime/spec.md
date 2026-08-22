## ADDED Requirements

### Requirement: Profile is bounded and source-aware

`presentation.designProfile()` SHALL return deterministic descriptive evidence
with a maximum of 256 entries per bounded collection. An imported profile SHALL
include `source.sourceBound: true` and the exact source revision SHA-256; a
source-free profile SHALL omit revision authority.

#### Scenario: Imported profile

- **WHEN** an Agent requests a profile from a trusted imported PPTX
- **THEN** the result includes the revision binding, design-language evidence,
  opaque summaries, and defensive candidate summaries without source bytes

#### Scenario: Source-free profile

- **WHEN** an Agent requests a profile from a newly authored Presentation
- **THEN** the result remains descriptive and reports that source-bound
  component candidates are unavailable

### Requirement: Profile cannot become a raw mutation surface

The profile SHALL NOT expose raw XML, package paths, relationship selectors, or
arbitrary attribute names, and profile inspection SHALL not change export bytes.

#### Scenario: No-op after profile inspection

- **WHEN** an imported PPTX is profiled and exported without another mutation
- **THEN** the output bytes equal the source bytes exactly
