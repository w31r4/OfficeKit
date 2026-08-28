## Context

The wire and native importer already preserve shape-tree order. The loss occurs after hydration: per-type JavaScript collections discard one global sequence, and export later emits fixed buckets. Imported source-bound export additionally requires all retained source elements to remain an ordered prefix, so arbitrary array movement cannot safely implement native reorder.

## Goals / Non-Goals

**Goals:**

- Preserve true direct-child order across every supported visual type.
- Make authored image backgrounds, scrims, editable text, charts, and foreground controls deterministic.
- Let inspection explain every element's stack position and reorder capability.
- Support bounded imported reorder without collateral serialization or unknown-content loss.
- Prove the complete workflow on source-free pages and the existing three complex PPTX references.

**Non-Goals:**

- No universal DOM/AST for every Office object.
- No XPath, raw part path, arbitrary XML patch, or implicit fallback.
- No claim that every third-party object can be reordered.
- No new template format or duplicate layer sidecar.

## Decisions

### 1. Slide owns the direct scene stack

`slide.elements` is the ordered direct-child collection. Type collections remain useful filtered indexes and keep their current add APIs. Every direct add registers once in the scene stack; group children continue using their existing ordered `children` collection. Deletion removes an element from both its type index and owner stack.

The stack is bottom-to-top, matching PresentationML shape-tree order and drawing behavior. Cross-type order methods operate only inside one owner and reject a target from another slide or group.

### 2. Ordering methods are shared, not copied per element type

A common layer helper implements capability reporting and the four order operations. Authored direct elements and authored group children are editable. Imported objects are marked source-bound during hydration and use a capability issued by the importer/codec. Unknown or unsafe imported objects remain inspectable and preserved but reject reorder.

### 3. Background image is an ordinary bottom-layer image

`slide.setBackgroundImage({ blob | dataUrl, fit, crop, ... })` creates or replaces one tagged full-slide image and moves it to index zero. It does not pretend that PresentationML supports a native bitmap background. The returned `ImageElement` remains editable, inspectable, and auditable. The background role is metadata in the JavaScript authoring model; export remains ordinary editable picture XML.

### 4. Imported reorder is an explicit source-bound mutation

The JavaScript exporter sends retained source elements in requested order only when each moved direct element has an issued reorder capability and all source elements remain present. The C# codec reopens the source SlidePart, verifies the full original sequence and element bindings, then moves existing native nodes in place. It changes only the target slide XML and rejects groups, animation targets, unresolved connectors, or other dependency-sensitive cases until those proofs exist.

Authored overlays on an imported slide remain a separate top-layer profile. Reordering source elements and adding overlays in one export is rejected; commit and reopen separates the mutations.

### 5. Acceptance reuses existing lossless assets

The existing `evals/pptx-lossless/manifest.v1.json` is the immutable source inventory. New layer evidence records source order, moved IDs, changed parts, masked equality, second import, and render results. A source-free photo/scrim/text page proves the visual composition that exposed the defect. The three third-party files prove detection and preservation; only files with an issued capability are mutated.

## Risks / Trade-offs

- [Code expects type-bucket ordering] -> Keep type collections intact and replace only order-sensitive consumers with the scene stack.
- [Imported node movement changes dependencies] -> Start from a narrow direct-child capability and have the codec reject timing, connector, group, or ambiguous identities.
- [Background image API implies native slide background] -> Document and inspect it as an editable picture role, not `p:bg`.
- [Visual review misses obstruction] -> Inspect stack order and run one host-rendered photo/scrim/text acceptance page in addition to layout checks.
