# Shape 3-D extrusion height

## ADDED Requirements

### Requirement: expose one bounded shape extrusion height

The system MUST expose a `shape3dExtrusionHeightEmu` native leaf only for an
ordinary shape with one direct child-free `a:sp3d` whose `extrusionH` is a
canonical non-negative bounded coordinate and whose other attributes, if
present, are standard direct 3-D attributes.

#### Scenario: project a direct extrusion height

- **WHEN** an imported shape contains a child-free direct `a:sp3d/@extrusionH`
  with a canonical bounded value
- **THEN** the projected shape exposes that value as
  `shape3dExtrusionHeightEmu` while depth, contour, material, and the rest of
  the 3-D state remain source-owned

#### Scenario: edit only the extrusion-height token

- **WHEN** a source-bound PPJ edit changes an issued extrusion-height leaf to a
  different canonical bounded coordinate
- **THEN** only `a:sp3d/@extrusionH` changes in the owning SlidePart and a
  subsequent projection reports the new value

#### Scenario: reject complex or malformed 3-D state

- **WHEN** the `sp3d` owner has child elements, unknown attributes, a missing,
  malformed, negative, noncanonical, out-of-range, or stale extrusion height
- **THEN** the system MUST keep this scalar source-owned or fail closed rather
  than partially reconstructing the 3-D graph
