# Change: six-sample programmable PPTX import

## Why

The earlier three-file proof established the source-bound edit model, but the
visual and structural variety in six newer samples can still expose import
omissions.  This change turns those real files into a repeatable, clean-room
evidence set without distributing the source artifacts.

## What changes

- Classify every visible top-level object in six real PPTX samples as typed,
  native-leaf, source-derived, or opaque-preserved.
- Preserve source revision hashes and design-profile evidence while allowing
  safe text, placement, and source-slide reuse operations.
- Add a deterministic evidence runner that proves no-op byte identity, second
  import, bounded package footprint, and task-local source protection.
- Record the samples, public provenance, and results without committing any
  copyrighted or licensed reference PPTX.

## Non-goals

This change does not claim arbitrary OOXML coverage, edit OLE internals, rebuild
unsupported topology, or provide Windows PowerPoint playback evidence.  It does
not change the Office wire protocol, PDF/Spreadsheet/Document codecs, template
assets, or presentation template format.
