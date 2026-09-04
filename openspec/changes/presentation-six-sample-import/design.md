# Design

## Projection boundary

The codec keeps three layers distinct:

```text
OPC source graph (exact bytes and relationships)
        ↓ bounded projection
typed model + native leaves + read-only opaque objects
        ↓ inspected capability
source-bound edit plan
```

Every top-level visible shape-tree child receives a source locator. A target is
one of `typed-editable`, `native-leaf-editable`, `source-derived-reusable`, or
`opaque-preserved`; the status describes the operation boundary, not whether the
object is visible. Opaque objects retain their original markup and relationship
closure. Their bounded descendant text is exposed read-only through
`nativeObject.inspectRecord()` so an Agent can understand a rich object without
turning it into a lossy table or shape.

## Corpus observations

The six samples exercise 1,694 visible top-level objects across 157 slides.
Most objects are already typed or have safe native leaves. The remaining
opaque records are concentrated in relationship/extension-sensitive MMS
objects and one rich Business Infographic table; these stay preserved rather
than being flattened. Imported tables may use a graphic-frame scale different
from their table grid, and covered merge cells may have no text leaf. The table
profile accepts those facts with a bounded scale ratio and refuses adding text
where no source leaf exists.

## Fidelity rule

An unchanged import exports the original package bytes exactly. A source-bound
edit must identify the current revision and target hash, then change only the
declared SlidePart or media payload. Every result is reimported before another
edit. Failure to locate a unique leaf, validate a dependency, or retain an
unmodelled shell is a refusal, not a fallback to a second authoring engine.
