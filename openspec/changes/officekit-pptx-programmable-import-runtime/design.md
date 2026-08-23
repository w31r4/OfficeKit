## Runtime model

The existing source package remains the lossless graph. The JavaScript
Presentation remains the semantic projection. A bounded Edit Plan remains the
only compiler mutation input. This change adds an accounting view, not another
model:

```text
trusted source revision
  -> source shape-tree entry
  -> semantic object and issued capabilities
  -> one primary classification record
```

The importer already retains one ordered source entry for every top-level
shape-tree object. Classification consumes those entries, the current
native-leaf capability, and component occurrence capabilities. It exposes no
raw XML, part path, relationship ID, or arbitrary selector.

## Classification precedence

An object receives one primary state in this order:

1. `typed-editable` when at least one public typed operation is explicitly
   source-backed, including semantic mutation, bounded text, accessibility,
   embedded Office package replacement, SVG leaf editing, or codec-proven
   deletion.
2. `native-leaf-editable` when no typed operation is issued but one or more
   current-revision native leaves are available.
3. `source-derived-reusable` when no edit operation is issued but a current
   component occurrence has codec-proven reuse permission.
4. `opaque-preserved` otherwise.

The record lists all available typed operations, native leaf kinds, and reuse
evidence even though only one primary state is chosen. Consumers must invoke
the corresponding existing API; changing a classification record has no
effect. Hashes and locators are regenerated after every import or resume.

## Completeness oracle

The independent evaluator opens the original ZIP and tokenizes each
`p:spTree`. It counts only direct visible children (`p:sp`, `p:pic`,
`p:graphicFrame`, `p:cxnSp`, `p:grpSp`, `p:contentPart`, and supported
compatibility wrappers), excluding shape-tree metadata. For every slide it
requires a one-to-one match with source-bound classification records by slide
and source shape-tree index. Duplicate locators, missing indices, unknown root
types, and extra classification records fail closed.

The runtime record count alone is not considered proof of completeness. The
three immutable benchmark manifests carry the independent oracle result.

## SVG leaves

SVG inspection remains byte-bounded and requires a base64 SVG source. Each
issued leaf contains the current SVG SHA-256, a stable leaf ID, a value hash,
and a finite operation kind. Edits splice only the token range that was issued:

- direct `fill` and `stroke` colors on one element;
- direct `opacity`, `fill-opacity`, or `stroke-opacity` in the closed interval
  `[0, 1]`;
- a local transform limited to a single existing translate/scale/rotate scalar
  tuple whose topology is unchanged.

CSS classes, `<style>`, inherited computed styles, paint servers, filters,
external resources, event attributes, scripts, and arbitrary transform-list
rewrites remain unsupported.

## Commit and gate discipline

Definition, runtime, public wiring, tests, real-sample evidence, Skill/Help,
and release gates are separate commits. Every functional commit runs the fast
gate; every completed milestone runs the Presentation slow segment. Full
repository and release gates run at most once per 24 hours until the release
candidate milestone.
