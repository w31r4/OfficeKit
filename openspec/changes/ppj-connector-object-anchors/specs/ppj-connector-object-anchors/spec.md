## Purpose

Make PPJ object-relative connector endpoints describe the actual compiled geometry and preserve their identities when a presentation is projected and edited again.

## ADDED Requirements

### Requirement: Object anchors determine connector endpoints

For final expanded non-connector objects, the compiler SHALL resolve `top`, `right`, `bottom`, `left` and `center` against the target's frame before its rotation and flips, then apply its ancestor group transforms. It SHALL express the resulting point in the connector's parent coordinates. The connector's frame SHALL NOT replace an object endpoint. Unsupported targets, missing IDs, invalid transforms and unrepresentable coordinates MUST fail with an identifiable diagnostic rather than fallback or clamping.

#### Scenario: Directed endpoints follow the target
- **WHEN** a left anchor targets a frame at `(750,300,60,40)` and a right anchor targets `(450,100,60,40)`
- **THEN** the compiled connector runs from `(750,320)` to `(510,120)` and retains its end arrow
- **AND** moving only the second frame down 40 makes the end `(510,160)`

#### Scenario: Group and rotated target
- **WHEN** an endpoint targets a flipped or rotated object inside a group with a distinct child frame
- **THEN** its side is transformed through that child space to the connector's space exactly once

### Requirement: Auto anchors are deterministic

An `auto` endpoint SHALL consider the four transformed side midpoints. The compiler SHALL select the candidate pair with minimum squared slide-space distance, keeping explicit or literal endpoints fixed. Equal-distance ties SHALL use top, right, bottom, left order for the start and then the end. Center SHALL be used only when explicitly requested.

#### Scenario: Auto-to-auto follows placement
- **WHEN** two separate rectangles are horizontally aligned and use auto endpoints
- **THEN** the shortest opposing side pair is selected, independently of element declaration order

### Requirement: Expansion preserves endpoint coordinate spaces

Component expansion SHALL rewrite local endpoint target IDs and transform literal endpoints with the same component placement as their sibling frames. Group descendants SHALL remain in the declared child space; an outer component placement SHALL transform the group's external frame once.

#### Scenario: Translated scaled component contains a connector
- **WHEN** a component containing shapes and literal/object endpoints is translated and scaled
- **THEN** the resulting endpoints agree with the expanded objects, including when a group with a nontrivial child frame is nested in the component

### Requirement: Endpoint identity survives round trip and editing

Export and projection SHALL retain supported object-anchor identity even without the embedded authored PPJ snapshot. Source-bound no-op SHALL preserve source bytes. Editing an anchor, replacing an object endpoint with a literal point, or moving its target SHALL update the connector consistently and preserve unrelated package parts. A PPJ frame anchor MUST NOT be represented as an arbitrary native geometry site index. Native site bindings and PPJ frame-anchor metadata SHALL remain distinguishable; unsupported imported bindings SHALL remain source-owned.

#### Scenario: Fresh projection preserves a frame anchor
- **WHEN** an authored file is stripped of its embedded PPJ snapshot and projected
- **THEN** its endpoints still identify their target objects and exact requested anchors
- **AND** a subsequent target move updates the connector, survives reprojection, and leaves other ZIP parts unchanged

#### Scenario: Remove object attachment
- **WHEN** an endpoint is replaced with explicit coordinates in a source-bound request rebuilt from the original projection
- **THEN** the matching attachment metadata is removed, the other endpoint is retained, and reprojection returns the explicit coordinates
