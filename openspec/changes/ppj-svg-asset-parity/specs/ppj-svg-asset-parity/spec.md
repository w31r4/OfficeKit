## ADDED Requirements

### Requirement: Paired SVG projection
An imported native picture with a safe raster/SVG fallback pair SHALL project
both members as separate content-addressed PPJ assets.

#### Scenario: Agent inspects an SVG-led imported page
- **WHEN** the picture projects to PPJ
- **THEN** `image.asset` names the raster fallback, `image.svgAsset` names the
  exact SVG, and the image nativeRef advertises `replaceSvg`

### Requirement: Source-bound SVG replacement
A capable `image.svgAsset` edit SHALL replace only the native SVG member while
preserving the raster fallback and picture topology.

#### Scenario: Replace one paired SVG
- **WHEN** an Agent supplies a declared `image/svg+xml` asset and changes only
  `image.svgAsset`
- **THEN** build retains the fallback, crop, frame, geometry, effects, and
  unrelated package content, and reimport recovers the new SVG hash

### Requirement: SVG topology fails closed
The compiler SHALL reject SVG-pair addition, removal, MIME mismatch, stale
nativeRef, or unissued replacement.

#### Scenario: Add SVG to an ordinary raster picture
- **WHEN** an Agent adds `svgAsset` to a picture without `replaceSvg`
- **THEN** build rejects the program instead of manufacturing a new fallback
  pair

### Requirement: Agent discoverability
The generated PPJ reference SHALL explain the raster fallback and SVG source
roles plus the required local-asset workflow.

#### Scenario: Agent edits imported SVG text
- **WHEN** an Agent reads the image reference
- **THEN** it knows to modify or regenerate the local SVG asset, update its
  declaration and `svgAsset`, build, reimport, and visually review the result
