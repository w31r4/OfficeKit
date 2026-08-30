# PPJ master and layout authoring

## Why

The native Presentation codec and the legacy object model already author one
canonical slide master, bounded layouts, direct backgrounds, text placeholders,
and slide-to-layout bindings. The PPJ registry classifies those APIs as
`ppj-state`, but the language has no corresponding fields. This makes a real
compiler capability invisible to Agents and lets the maintenance gate report a
false positive.

## What changes

- Add bounded `design.masters[]`, `design.layouts[]`, and `pages[].layout` state.
- Author one canonical master, `blank`/`title`/`titleOnly`/`obj` layouts, direct
  backgrounds, bounded master paragraph defaults, and direct-frame text
  placeholders.
- Preserve a third-party slide's source layout identity in projected PPJ while
  keeping imported master/layout topology read-only.
- Require every legacy API classified as `ppj-state` to name its concrete PPJ
  path, so ownership cannot be mistaken for implemented mapping again.

## What does not change

- PPJ does not expose raw Master/Layout OOXML or relationship IDs.
- Source-free authoring does not gain multiple masters, arbitrary template
  graphs, non-text placeholders, or inherited placeholder geometry.
- Imported Master/Layout topology and semantic mutation remain source-owned and
  fail closed.
