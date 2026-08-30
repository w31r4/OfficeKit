# Imported PPTX and nativeRef

Use this route for a third-party PPTX whose unknown package content must survive
an edit.

```bash
officekit ppj import input.pptx -o deck.ppj
officekit ppj inspect deck.ppj --query "visible title" --json
officekit ppj check deck.ppj --json
officekit ppj build deck.ppj -o edited.pptx --json
```

Import copies the source into a content-addressed read-only asset and binds its
SHA-256. Every visible object must appear as a typed element or an `opaque`
element with `nativeRef`, location, summary, and issued capabilities. Unknown
OOXML remains in the source package and does not enter the model context.

## Edit boundary

- Edit ordinary PPJ fields for fully modelled objects.
- For an opaque or partially modelled object, edit only fields explicitly
  listed by its `nativeRef` capability.
- Keep source revision, target hash, object identity, and page identity intact.
- Re-import after build and locate the same stable IDs again.

No-op build must return the source bytes exactly. A supported edit may change
only the target part and necessary dependencies. Unrelated parts,
relationships, master/layout/theme state, unknown timing, OLE, SmartArt, and
other opaque topology must remain stable.

Some imported shapes with a strict direct embedded `a:blipFill` and a custom
geometry that is not yet semantically decoded may expose a source-bound
`imageFill` capability. This is a bounded frame-edit path: the existing image
relationship, crop, and custom geometry remain source-owned while the shape's
position or size changes. Image replacement, fill conversion, custom-path
rewriting, and any other image-fill graph stay opaque and fail closed.

Stale hash, ambiguous target, unsupported field, unsafe relationship change,
cross-object mutation, or topology rewrite must fail. Do not patch raw OOXML,
replace the whole slide with an image, flatten the deck, or rebuild it through
an authored route to make the request appear successful.

## Source continuation

A source page or component may be reused only when import issues a reuse
capability. The reused object remains source-derived and must be re-importable.
New PPJ objects may be composed around it without changing the unknown source
subgraph.

OfficeKit-authored PPTX is different: if a valid embedded program exists,
import restores it exactly. If an external application changed the native file
but left the program, the embedded PPJ remains authoritative. Build a new file;
never overwrite or silently merge native drift.

Use `--task` only when immutable revisions and resume evidence are useful. Task
state does not weaken source checks or restore process memory.
