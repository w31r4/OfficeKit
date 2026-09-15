## Why

The PPJ theme accent field already edits direct RGB `accent3`, while the backlog still leaves the direct `a:alpha` child source-owned. A bounded RGBA owner closes one concrete F-15 field without rewriting the remaining theme graph.

## What Changes

- Preserve a canonical imported `accent3` RGBA value when the shared ThemePart has a direct RGB leaf with one canonical `a:alpha` child.
- Reuse the existing `setThemeAccent3Color` capability for `accentColors.accent3`; the field owns the RGB and alpha pair together.
- On source-preserving export, patch only `a:accent3Color/a:srgbClr/@val` and its existing `a:alpha/@val`.
- Reject missing/non-canonical alpha, alpha-presence changes, deletion, combined fields, transformed colors, and capability tampering.

## Capabilities

### Modified Capabilities

- `presentation-theme-accent3-color`: the existing source-bound field now covers a strict direct RGBA accent3 owner in addition to direct RGB.

## Impact

The PPJ importer/projector, source-bound theme writer, registry/coverage docs, and focused native regression gain one bounded RGBA source owner. The existing repeated `PresentationThemeArtifact.accent_rgb` wire field and protocol version are reused.
