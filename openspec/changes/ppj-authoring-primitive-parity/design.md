## Context

The generated PPJ language manual currently has 2,202 lines and the compared
finite DSL manual has 1,886. PPJ declares 16 chart families and 176 preset
geometries. The comparison is therefore not a contest of document length or
top-level type count. The useful question is whether an Agent can express a
common visual intent concisely and deterministically.

## Decisions

### 1. Classify before adding

Every observation is classified as one of:

- semantic gap: the intended output cannot be expressed;
- convenience gap: the output exists but requires avoidable expansion;
- intentional constraint: PPJ rejects nondeterministic or unrecoverable state;
- already covered: current schema/compiler owns the semantics under another
  finite representation.

### 2. Keep line state unified

PPJ `connector` already accepts free points or element anchors, straight,
elbow or curved routing, stroke and arrowheads. It remains the one public line
primitive. A second `line` element would create two identities for the same
native object without adding output power.

### 3. Table inheritance expands before native lowering

Table-level cell styles are finite declarative sugar. The compiler resolves a
baseline, cycling body-row styles, first/last row and first/last column styles,
then the explicit cell. `rowOverColumn` selects only row/column conflict order.
The expanded native table remains ordinary editable cells; no style program is
embedded into DrawingML.

### 4. Vector Sankey remains deterministic

`right` alignment uses reverse topological depth. A node-name color map
overrides the finite palette only for declared nodes. Unknown names, duplicate
categories, cycles and non-conserving internal flows continue to fail before
native writing.

## Non-Goals

- Fetching remote images or icon libraries at compile time.
- Pretending inline LaTeX is ordinary DrawingML text.
- Adding raw native chart XML, arbitrary expressions or an unbounded style
  cascade.
- Expanding every chart option in one change.

## Follow-up boundary

Named icons need a pinned offline catalog plus license and deterministic asset
lowering. Formula text needs an explicit editable/native-vs-vector decision.
Native pie geometry, axis reversal/arrows/line styles, bubble sizing and richer
combo topology require chart wire and writer work. Those are true follow-ups,
not hidden inside this no-wire batch.
