## Why

PPJ can embed a local SVG asset, but an Agent still has to discover a file,
manage an asset declaration and preserve its hash for every common interface
icon. A mature finite presentation DSL exposes a named icon directly. The
missing abstraction is therefore not arbitrary remote imagery; it is a pinned,
offline and licensed icon vocabulary that the compiler can lower without
network access or JavaScript preprocessing.

## What Changes

- Add a typed `icon` PPJ element with a finite `iconName` and ordinary native
  frame, paint, accessibility and transform state.
- Pin the Font Awesome Free catalog used by the language and generate a compact
  checked-in compiler resource from the public packages.
- Lower named icons deterministically to editable DrawingML custom geometry.
- Preserve exact icon intent through the embedded PPJ snapshot used by
  OfficeKit-authored presentations.
- Reject unknown names before PPTX writing and keep imported arbitrary vector
  shapes classified as shapes rather than guessing icon identities.
- Publish the catalog license, version and brand-use boundary in the Agent
  manual and third-party inventory.

## Capabilities

### New Capabilities

- `ppj-named-icon-primitive`: Offline named icons that compile into native,
  editable Presentation geometry without remote fetches or asset sidecars.

### Modified Capabilities

None. The PPJ schema ID and Office wire version remain unchanged.

## Impact

The PPJ schema and typed C# model, generated icon catalog, authored compiler,
generated language reference, Presentation guidance, third-party notices,
coverage and one existing authored PPJ contract are affected.
