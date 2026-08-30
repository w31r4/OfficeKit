# PPJ solid-background opacity

## Why

PPJ solid fills already accept alpha and slide backgrounds already compile to
native `p:bg`, but the authored compiler rejects any translucent solid
background. This leaves one basic layer primitive inconsistent with the rest
of the language and weakens image-overlay composition.

## What Changes

- Add presence-aware opacity to the existing native presentation background.
- Read and write one direct DrawingML `a:alpha` transform on solid RGB or theme
  colors.
- Compile and project PPJ solid-background opacity without introducing a new
  element or overlay convention.

## Impact

- Additive Office wire-v2 field only; no wire-version change.
- One existing native background contract covers author, import, edit and PPJ
  projection.
