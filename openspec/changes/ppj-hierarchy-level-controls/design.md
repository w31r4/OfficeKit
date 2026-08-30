# Context

Treemap and sunburst are compiler-owned editable DrawingML groups. Their PPJ
data already carries a complete, bounded tree and proves parent totals. A
display-depth control belongs to that semantic series, not to individual
generated shapes.

# Decisions

## 1. Add one bounded series field

`levels` is an optional integer on the sole treemap or sunburst series. Level 1
means roots only; level 2 includes roots and their direct children. Treemap
accepts 1..8 and sunburst accepts 1..6, matching their existing maximum tree
depths. Other chart families reject the field.

## 2. Preserve full data semantics

Validation continues to inspect all nodes, parent links, cycles and direct
child totals. `levels` changes only the compiled view. When the declared value
exceeds the actual depth, the compiler renders the complete available tree.

## 3. Reflow rather than hide after layout

A treemap node at the last visible level becomes the visible leaf and receives
the full rectangle assigned to that branch. A sunburst divides the available
radius by the visible level count, so fewer levels create wider readable rings.
Generated descendants beyond the limit do not enter the native group.

## 4. Keep recovery honest

Embedded PPJ restores the full hierarchy and its `levels` intent. Without the
snapshot, import returns only the editable shapes that actually exist and does
not invent hidden semantic nodes.

# Lean verification

Extend the existing comprehensive authored-PPJ test by limiting its existing
treemap and sunburst examples. Assert that hidden descendants are absent from
the native group while embedded recovery retains `levels`. Do not add a new
fixture, hierarchy matrix or benchmark.
