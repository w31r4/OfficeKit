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

- Edit ordinary PPJ fields only when the typed object has the corresponding
  issued capability, such as `replaceText`, `setFrame`, `replaceImage`,
  `setChartTitle`, or `setChartData`.
- For a precise imported scalar, use an issued `nativeRef.leaves[]` entry. It
  contains an opaque leaf ID, closed `kind`, source-value `expectedHash`, and
  human-facing `value`. Change only `value`.
- Never invent a leaf, copy one to another object, alter its ID/kind/hash, or
  convert a typed operation such as image replacement into a scalar leaf.
- Keep source revision, target hash, object identity, and page identity intact.
- Re-import after build and locate the same stable IDs again.

Example imported run-size edit:

```json
{
  "id": "nl_8b1f…",
  "kind": "fontSizePoints",
  "expectedHash": "<sha256 of the exact native old value>",
  "value": 20
}
```

The value uses a human unit where PPJ defines one: points for font size,
degrees for rotation, booleans for flips, and `#RRGGBB` for direct RGB. EMU
leaves remain explicitly named `*Emu`; they exist for exact source surgery,
not for ordinary authored layout.

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

An imported shape may also expose `textEditable` while its fill or effect graph
is intentionally opaque (for example, a native gradient banner). A plain
text-only replacement is safe when the original paragraph/run topology is
kept; the native fill, effects, geometry, and relationships remain byte-owned
by the source. Style, frame, name, and topology changes still fail closed.

Source-bound connectors may expose the same bounded placement surface when
their direct frame is a legal horizontal or vertical line with one zero
extent. Moving the frame leaves endpoint bindings, line geometry, and unknown
extension children untouched. Zero-by-zero, negative, missing, or ambiguous
connector frames remain read-only; this is not permission to rewrite connector
topology.

Source-bound opaque pictures may likewise retain a bounded negative left/top
offset when the picture frame and unique image relationship are proven safe.
This supports intentional edge bleed and crop layouts; the image payload,
effects, crop, and relationship remain source-owned.

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
