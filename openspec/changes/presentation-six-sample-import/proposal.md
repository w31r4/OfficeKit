# Six-sample programmable PPTX import

## Why

The three-file lossless slice proved that OfficeKit can keep an imported
presentation intact while editing a bounded target. The next useful boundary is
broader real-world variety: two NASA technical decks and four SlidesCarnival
decks contain dense groups, pictures, connectors, tables, charts, EMF/SVG
assets, notes, and extension-heavy slide markup that are not represented by the
small internal fixtures.

## What changes

- Freeze six locally downloaded PPTX files as ignored, read-only evaluation
  inputs and record their hashes and package inventories without redistributing
  the source bytes.
- Make imported objects discoverable through stable `importObject` locators,
  classification, dependency summaries, and source revision hashes.
- Expand only bounded, source-preserving profiles that are demonstrated by the
  corpus: richer pictures, scaled/merged DrawingML tables, and readable text
  on opaque objects. Keep unsupported rich cells, connectors, OLE, WPS, and
  extension-heavy topology opaque when an edit cannot be proven safe.
- Record no-op byte identity and real local table/image edits with second-import
  and target-only SlidePart evidence.
- Teach the Presentations Skill that import is `inspect → classify → resolve →
  edit or stop → export → reimport → review`; a visible but opaque object is
  still inspectable and must not be flattened.

## Non-goals

This change does not claim complete OOXML semantics, does not add raw XML or
XPath access, does not rebuild imported slides, and does not include the
external PPTX files in source or npm packages. Windows PowerPoint playback and
full visual host acceptance remain separate evidence.
