# PPJ text highlight

## Why

OfficeKit already reads, writes and source-patches direct DrawingML run
highlights, but PPJ cannot declare the same state. This leaves an implemented
presentation primitive invisible to Agents and forces authored PPJ to omit a
common editorial emphasis.

## What Changes

- Add `textStyle.highlight` using the existing PPJ color value.
- Compile direct authored run and default-run highlights.
- Project safe direct RGB highlights from imported presentations.
- Keep imported theme highlights available through the existing capability-
  issued native leaf until source theme tokens have a lossless PPJ mapping.
- Document highlight as text emphasis, not a generic badge or card background.

## Impact

- No protocol change: wire v2 already contains RGB and theme highlight fields.
- No wider source-bound text-style mutation: imported edits continue through
  exact `fontHighlightRgb` or `fontHighlightScheme` leaves.
- One existing integrated PPJ contract gains focused assertions.
