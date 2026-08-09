# PowerPoint Slide Transitions

This `p:transition` contract covers the complete ECMA-376 base-namespace
effect vocabulary for Agent workflows. It controls the transition between
slides; it is not a PowerPoint timing or animation engine, and a static PNG/PDF
render cannot prove slideshow playback.

## Source-Free Authoring

```ts
const opening = presentation.slides.add({
  name: "Opening",
  transition: {
    effect: "fade",
    throughBlack: true,
    speed: "medium",
    advanceOnClick: true,
    advanceAfterMs: 4_000,
  },
});

const decision = presentation.slides.add({ name: "Decision" });
decision.setTransition({
  effect: "split",
  orientation: "horizontal",
  direction: "in",
  speed: "fast",
  advanceOnClick: false,
});
```

The 21 base effects use these exact semantic profiles:

| Effects | Effect-specific fields | Default |
| --- | --- | --- |
| `blinds`, `checker`, `comb`, `randomBar` | `orientation`: `"horizontal"` or `"vertical"` | `horizontal` |
| `circle`, `diamond`, `dissolve`, `newsflash`, `plus`, `random`, `wedge` | none | — |
| `cover`, `pull` | `direction`: cardinal or corner | `left` |
| `cut`, `fade` | optional boolean `throughBlack` | omitted |
| `push`, `wipe` | `direction`: `"left"`, `"up"`, `"right"`, or `"down"` | `left` |
| `split` | `orientation` plus `direction`: `"in"` or `"out"` | `vertical`, `out` |
| `strips` | `direction`: `"leftUp"`, `"rightUp"`, `"leftDown"`, or `"rightDown"` | `rightDown` |
| `wheel` | `spokes`: integer `1..8` | `1` |
| `zoom` | `direction`: `"in"` or `"out"` | `in` |

Every profile also accepts:

| Field | Supported values | Default |
| --- | --- | --- |
| `speed` | `"slow"`, `"medium"`, `"fast"` | `"medium"` |
| `advanceOnClick` | boolean | `true` |
| `advanceAfterMs` | integer `0..86400000` | omitted |

Fields that do not belong to the chosen effect are rejected. The codec writes
one direct `p:transition` with explicit `spd` and `advClick`, an optional
numeric `advTm`, and exactly one canonical base effect child.
`clearTransition()` removes that direct element.

## Inspect, Resolve, And Edit

```ts
const inspection = await presentation.inspect({ kind: "slide,transition" });
const transition = presentation.resolve(`${opening.id}/transition`);

if (transition.capability.editable) {
  transition.set({ effect: "wheel", spokes: 6, speed: "slow" });
}
```

Each slide emits a stable `${slide.id}/transition` inspect record, even when it
is not configured. Its capability is evidence for routing, not permission to
edit an arbitrary package:

```ts
{
  sourceBound: boolean,
  partPresent: boolean,
  editable: boolean,
  addable: boolean,
}
```

Source-free slides are addable and editable. A source-bound slide is editable
only when its existing transition fits this exact profile. An imported slide
with no transition is addable only when `capability.addable` is true: that
proves its root contains only `p:cSld` plus optional `p:clrMapOvr`, with no
`p:transition`, `p:timing`, or extension leaf. Adding the canonical direct
leaf is then a one-SlidePart operation, not a general animation patch.

## Imported And Clone Boundary

The C# Open XML SDK codec accepts one direct `p:transition` only when its
attributes and single child exactly match one of the profiles above.
`p:timing`, `p14:dur`, sound actions, extension lists, extra or unknown
attributes, multiple children, Office-version effect extensions, malformed
timers, and any broader timing/animation graph remain opaque. They are
preserved byte-for-byte when unrelated supported edits occur, and
`setTransition()` or `clearTransition()` rejects them.

The strict imported `slide.duplicate()` profile may copy one unchanged
canonical direct base transition with its SlidePart. It does not copy or
interpret a timing tree, sound relationship, or extension graph. Re-import
after export before making another semantic change.

## Source-Bound Transaction

For an Agent request to replace one existing transition, use the shipped
`examples/officekit-transition-edit-workflow.mjs` transaction rather than a
raw XML patch:

```bash
officekit run examples/officekit-transition-edit-workflow.mjs \
  input.pptx output.pptx audit.json \
  "Decision" \
  '{"effect":"fade","throughBlack":true,"speed":"medium","advanceOnClick":true}' \
  '{"effect":"split","orientation":"horizontal","direction":"in","speed":"fast","advanceOnClick":false}'
```

The fourth argument is one unique imported slide name. The two JSON objects
are the complete expected source state and desired replacement state; defaults
use the same public transition contract above. The workflow admits only an
existing source-bound canonical direct transition with `partPresent: true` and
`editable: true`. It does not add a transition to an absent source or clear an
existing one.

Before publication it binds source bytes, proves the expected source state and
replacement state, writes a temporary output, and requires unchanged package
topology plus byte-identical non-target parts.
Exactly the selected SlidePart must differ. It reimports the replacement
semantics, preserves every slide's
non-transition model snapshot, checks static SVG render hashes, runs model
verification, and writes a source/output-bound audit without replacing an
existing output path. A missing or duplicate slide name, stale expected
transition, opaque/timing/sound/extension graph, unexpected package change,
no-op replacement, or output collision fails closed. Static proof still does
not certify real PowerPoint playback.

## Verification

After a mutation, export, re-import, and inspect the transition record again.
For a source-bound edit, verify that the intended SlidePart transition changed
and unrelated package parts stayed within the operation's declared scope.
LibreOffice/Poppler visual QA can prove that visible static slide content did
not regress, but cannot certify transition playback itself. Use a real
PowerPoint/native-host slideshow QA lane when playback timing or host-specific
effects matter.
