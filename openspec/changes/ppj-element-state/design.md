## Context

PresentationML stores object visibility on `p:cNvPr@hidden`. Editing locks are
DrawingML children whose element type differs for shapes, pictures, connectors,
graphic frames, and groups. Existing OfficeKit writers also use a few baseline
locks for native correctness. PPJ exposes one optional boolean, so the compiler
must define a canonical cross-type profile rather than leak that native
variation into the language.

## Goals / Non-Goals

**Goals:**

- Make every source-free typed PPJ element honor explicit `hidden` and
  `locked` state.
- Give `locked: true` a predictable Agent-facing meaning: the object cannot be
  selected, moved, resized, rotated where supported, or edited through its
  type-specific geometry/text/crop controls.
- Make `locked: false` retain only OfficeKit's ordinary baseline locks.
- Recover recognized state from third-party PPTX and allow a hash-bound local
  state edit without rebuilding the object.
- Fail closed on partial or private native lock profiles.

**Non-Goals:**

- A public per-flag OOXML lock API, selection-pane automation, password
  protection, slide visibility, document protection, or host permission model.
- Rewriting an imported object's unknown lock extension records.

## Decisions

### 1. Optional wire presence is authoritative

`PresentationElement` receives optional `hidden` and `locked` fields. Absence
means PPJ did not declare the state. Explicit `false` remains distinguishable
from absence through NativeAOT validation and source-bound diffing.

### 2. One canonical full-lock profile per native object kind

The native writer maps `locked: true` to all applicable standard edit locks for
that object kind. It maps `locked: false` to the writer's normal baseline:
chart/table graphic frames retain `noGrouping`; media posters retain
`noChangeAspect`; ordinary shapes, images, connectors, and groups retain no
additional lock. This makes the public boolean stable without exposing native
flag names.

### 3. Imported state is editable only when exact

Import recognizes an unlocked baseline and OfficeKit's complete locked profile.
Any other standard-flag combination, extension child, or unrecognized native
state is preserved in the source package but does not issue `setState`.
Visibility remains independently recognized when `cNvPr` itself has no unknown
state-bearing structure.

### 4. State mutation is a token-bounded local object change

The PPJ source-bound compiler requires `setState`, changes only the optional
wire fields, and the PPTX exporter re-proves the exact native profile before
editing the existing non-visual node. Content, relationships, z-order, IDs, and
unknown descendants remain untouched.

## Risks / Trade-offs

- [PowerPoint may display a fully locked object differently across hosts] → Use
  standard DrawingML flags, validate package structure, and avoid claiming host
  UI behavior without host evidence.
- [A boolean can erase meaningful partial locks] → Never issue the capability
  for a partial profile.
- [Baseline native locks look like PPJ locks] → Normalize baseline locks per
  content type before classifying the public state.
- [Nested group state is overlooked] → Apply the same recursive element
  projection and owner-local source proof used by existing group edits.

## Migration Plan

Add fields, codec, compiler/projector support, guidance, and one focused
round-trip assertion. Existing PPJ without the fields produces identical
output. Rollback removes the additive fields and restores the explicit
compiler rejection.

## Open Questions

None. A future typed lock profile would require a new PPJ schema version.

