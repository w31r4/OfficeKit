## Context

PPJ deliberately requires remote images to be downloaded, hashed and declared
as assets. Named icons are different: a finite icon library behaves like a
font or preset-geometry catalog. Requiring every Agent to rediscover and copy
the same SVG file adds ceremony without improving provenance.

The compiler already writes DrawingML custom geometry with line, quadratic,
cubic and arc commands. A generated catalog can therefore supply normalized
paths to the NativeAOT compiler without adding a network provider or a Node
runtime dependency.

## Decisions

### 1. One pinned offline catalog

The public names use Font Awesome Free prefixes:

- `fas:` for free solid icons;
- `far:` for free regular icons;
- `fab:` for free brand icons.

The exact package versions are development inputs to a deterministic generator.
The distributable contains one compact generated catalog and the required
license notices, not the npm packages or thousands of loose SVG files.

### 2. Icon is a typed visual element

An icon has a stable element ID, frame, `iconName`, paint, transform,
accessibility and normal z-order. It does not accept raw SVG, arbitrary path
data, a remote URL or an icon-library expression.

The catalog's view box is fitted inside the declared frame while preserving
aspect ratio. `fill`, `stroke`, opacity and ordinary element transforms remain
Agent-controlled. Multiple contours remain bounded subpaths inside the one
native custom-geometry shape and one PPJ identity.

### 3. Native vector lowering, not an image shortcut

The C# compiler parses only the bounded normalized command vocabulary emitted
by the generator and produces DrawingML custom geometry. The result remains
editable, recolorable and inspectable in PowerPoint. No dynamic SVG asset or
relationship is created.

If the upstream path cannot be normalized within the supported finite command
set, catalog generation fails. Runtime compilation never falls back to a
raster image or remote fetch.

### 4. Recovery is explicit, inference is not

OfficeKit-authored PPTX embeds the PPJ snapshot, so reimport restores the exact
`iconName`. If another application removes the snapshot, the native geometry
is projected as an ordinary custom shape. Third-party imported geometry is
never heuristically relabeled as a named icon.

### 5. Brand use remains semantically constrained

Brand icons are available because they are useful for identifying products,
companies and services. The Skill states that they must not imply sponsorship,
endorsement or ownership. Compilation proves identity and license provenance;
it cannot prove trademark context.

## Rejected alternatives

- Runtime URL lookup: nondeterministic, network-dependent and difficult to
  recover.
- Node-side icon expansion: violates the direct PPJ-bytes-to-C# boundary.
- Loose SVG asset expansion: repeats boilerplate and loses the named intent.
- Icon-font glyphs: font availability and glyph mapping are host-dependent.
- A generic raw-path PPJ primitive: exposes an unbounded low-level language and
  duplicates custom geometry internals.

## Lean verification

Extend one existing authored PPJ contract with one named icon, inspect its
native custom geometry, reimport the embedded PPJ and reject one unknown name.
Run the catalog generator in check mode, the Presentation Skill maintainer and
strict OpenSpec validation. Do not add an icon matrix, image snapshots or a
new test file.
