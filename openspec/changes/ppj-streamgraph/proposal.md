# Change: Add authored PPJ streamgraphs

## Why

PPJ can author native area charts, but it cannot express the centered flowing
bands used to compare changing composition over an ordered domain. Rebuilding
that visual manually from custom shapes loses data semantics and stable IDs.

## What changes

- Extend area-chart `style.stacking` with `stream`.
- Compile the bounded stream profile into editable DrawingML paths and text.
- Restore the exact PPJ program through the embedded authored-program snapshot.
- Teach the Presentation Skill when a streamgraph is appropriate and when it
  is less truthful than an ordinary area or line chart.

## What does not change

- No new wire operation or public chart type is added.
- Third-party arbitrary path groups are not inferred as streamgraphs.
- Imported native ChartParts do not receive a fake stream capability because
  DrawingML has no native centered-stream stacking mode.

