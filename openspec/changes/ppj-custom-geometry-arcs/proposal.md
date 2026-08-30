# PPJ custom-geometry arcs

## Why

The native codec and wire contract already preserve literal DrawingML arc
commands, but authored PPJ exposes only lines and Bezier segments. Agents must
therefore approximate circles, rings and organic masks manually, and a
third-party literal arc cannot remain typed when projected into PPJ.

## What changes

- Add one bounded `arcTo` command to the PPJ custom-path vocabulary.
- Express radii in view-box units and angles in degrees rather than exposing
  DrawingML integer-angle units.
- Compile to and project from native editable `a:arcTo` commands.
- Share the command between authored shapes and custom image masks.

## What does not change

- PPJ does not accept raw SVG path strings, raw OOXML or formula references.
- Adjustment guides, handles, connection sites and source-owned custom-path
  topology remain outside authored PPJ.
- Imported formula-driven or otherwise irregular geometry remains opaque.

