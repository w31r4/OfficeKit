## Implementation evidence

- Pinned build inputs: Font Awesome Free solid, regular and brands `7.3.1`,
  plus `svgpath 2.6.0` for deterministic path normalization.
- Generated catalog: `2,163` names, exact ordinal order, one checked-in
  `office-kit/ppj-icon-catalog/v1` resource, and only explicit absolute
  `M`/`L`/`C`/`Z` commands under the 512-command runtime bound.
- Native lowering: PPJ `icon` becomes one aspect-preserving editable DrawingML
  custom shape. It adds no media asset, relationship, icon font, remote fetch,
  SVG runtime or raster fallback.
- Source-bound continuation: a typed icon overlay compiles on the eligible
  imported page and changes only the target page scope already permitted by
  the source-bound authored-overlay contract.
- Recovery boundary: matching OfficeKit snapshots recover exact `iconName`;
  ordinary projection reports the native result as a shape with `nativeRef`
  and never infers a catalog identity.

## Lean verification

- `PpjSourceBoundProgramReusesOneProvenSlide`: passed once after the native
  semantic-oracle provenance mask was corrected.
- `PpjV1CompilesCanonicalPresentationProgramDeterministically`: passed once;
  it covers known-name compile, unknown-name rejection, native geometry,
  paint, accessibility, determinism and exact embedded PPJ recovery.
- `npm run ppj:icons:check`: `2,163` icons current.
- Presentation Skill maintainer: passed with `151` Help APIs, `73` native
  leaves and `13` host-only operations.
- `npx openspec validate ppj-named-icon-primitive --strict`: passed.
- `npm run proto:check`: passed; wire version remains `2`.
- `npm pack --dry-run --json`: the `2,090,110`-byte generated catalog is
  present in the package; no Font Awesome package or loose SVG tree is shipped.

No per-icon matrix, screenshot fixture, full `npm test`, NativeAOT release
link, or host PowerPoint visual run was added for this bounded language slice.
