## Purpose

Retain each custom-shape path's own coordinate extents, including native default axes, through PPJ authoring and source editing.

## ADDED Requirements

### Requirement: Per-path extents preserve defaults and differences

Custom shape paths SHALL accept optional viewport requiring nonnegative width and height in existing path units. Omission SHALL inherit geometry.viewBox extents; zero SHALL select the native default independently per axis. Positive, heterogeneous and zero extents SHALL survive export and fresh projection as equivalent effective native extents. Redundant overrides SHALL normalize to omission. Issued numeric path-width/height leaves SHALL retain values through the native maximum 2147483647; other leaf kinds SHALL retain their existing numeric bounds. Negative/out-of-budget extents and mask/clip viewport overrides MUST reject.

#### Scenario: Mixed default axes
- **WHEN** paths contain inherited extents, independent positive extents and zero on either or both axes
- **THEN** export and fresh projection SHALL preserve each effective native width/height and command

### Requirement: Source viewport overrides support lifecycle edits

Existing paths authority SHALL permit adding, changing and removing viewport overrides. Removal SHALL restore inheritance from the current geometry baseline. Source no-op SHALL preserve bytes. Edits SHALL preserve path order, commands, flags, other geometry state, text/frame and non-target package members, with canonical projection allowed to refactor baseline and overrides.

#### Scenario: Modify or remove an override
- **WHEN** separate original-source requests alter a viewport or remove it
- **THEN** fresh projection SHALL compile to the requested native extents and preserve unrelated content

### Requirement: Preview reports viewport limits

Preview SHALL diagnose per-path viewport semantics it cannot render.

#### Scenario: Viewport preview
- **WHEN** a custom path has a viewport override
- **THEN** unsupported viewport semantics SHALL NOT be reported as fully rendered
