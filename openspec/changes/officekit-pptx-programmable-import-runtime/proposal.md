## Why

OfficeKit can already preserve a trusted PPTX source package, perform exact
no-op export, edit a bounded set of semantic and native leaves, reuse source
slides, and continue an OfficeKit-authored task. Those capabilities are still
presented as separate features. An Agent cannot yet ask one deterministic
question that accounts for every visible source object, nor can it safely edit
the non-text styling inside the McKinsey sample's eight full-slide SVG assets.

The next runtime milestone is therefore not broader OOXML reconstruction. It is
to make the three real benchmark decks explicit, source-bound program states:
every visible top-level object is accounted for, every permission names the
operation it proves, and everything else remains opaque and byte-preserved.

## What Changes

- Extend `presentation.inspect()` with one source-bound classification record
  for every visible top-level object in a trusted imported PPTX.
- Assign exactly one primary state: `typed-editable`,
  `native-leaf-editable`, `source-derived-reusable`, or `opaque-preserved`.
- Report the exact issued operations, source revision, stable object locator,
  source hashes, and a bounded dependency summary. Classification itself never
  grants authority.
- Add an independent package oracle that compares raw slide shape-tree roots
  with the imported classification surface, so importer omissions cannot be
  hidden by self-reported model counts.
- Add source-bound safe SVG leaves for direct fill, stroke, opacity, and a
  bounded local transform profile. Continue to reject active content, external
  references, DTD/entity use, `foreignObject`, stylesheet topology, and
  ambiguous inherited styling.
- Extend source-derived pages only through existing typed operations and issued
  leaves. Unknown native subgraphs remain attached to the source graph and are
  never rebuilt to make an edit succeed.
- Integrate the independent intent matrix and Codex acceptance harness in a
  later evidence milestone without coupling evaluator code to runtime code.

## Non-Goals

- No universal Office AST, raw XPath/XML editing, second PPTX writer, or HTML
  conversion path.
- No promise that all OOXML vocabulary becomes semantic or editable.
- No mutation permission inferred from visual similarity or a design profile.
- No Windows PowerPoint work or acceptance gate in this portable runtime
  change; desktop host acceptance remains a separately scheduled lane.
