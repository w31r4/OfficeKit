# PPJ bullet color tokens and alpha

## Why

PPJ bullets use the normal color type, but authored paragraphs accept only an
opaque literal RGB string. Deck-local design tokens and alpha-bearing colors
are rejected even though other text paint resolves through the color catalog.

## What Changes

- Resolve bullet colors through the authored PPJ color catalog.
- Add presence-aware opacity to the bounded native bullet color.
- Import and project one direct DrawingML `a:alpha` child.

## Impact

- Additive Office wire-v2 field only.
- No new PPJ schema field or public JavaScript authoring API.
