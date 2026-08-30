## Implementation evidence

- PPJ rich-text runs accept exactly one of literal `text` or a typed
  `{ formula: { syntax: "latex", source } }`.
- The clean-room C# reader implements a finite, non-executable grammar under
  4,096 source characters, 512 tokens, 32 nesting levels and 2,048 AST nodes.
  It has no macro, environment, package, condition, counter, file, network or
  raw-OMML escape hatch.
- Additive protocol-v2 messages carry the finite math AST. NativeAOT writes
  editable `a14:m` / OMML and re-reads the canonical graph for semantic proof.
- Exact LaTeX is recovered from the embedded PPJ snapshot. Ordinary
  third-party OMML is preserved and never reverse-translated into guessed
  source or granted a mutation capability.
- Open XML SDK 3.3 models `a14:m` as a leaf despite the native child graph.
  Package validation ignores only that exact SDK child error after every
  affected formula independently passes OfficeKit's canonical OMML validator.

## Lean verification

- `PpjV1CompilesCanonicalPresentationProgramDeterministically`: passed once.
  The existing comprehensive case now covers mixed prose/formula order,
  integral with scripts, fraction, radical, Greek symbols, native OMML,
  deterministic output, exact embedded PPJ recovery and one unsupported
  environment rejection.
- `dotnet build` for `OfficeKit.Codec`: passed before the focused contract.
- `npm run proto:check`: passed after the generated protocol output was
  committed; Office wire remains version `2`.
- Presentation Skill maintainer: passed with `151` Help APIs, `73` native
  leaves and `13` host-only operations.
- `npx openspec validate ppj-inline-formula-primitive --strict`: passed.
- `git diff --check`: passed.

No formula corpus, effect matrix, screenshot snapshots, full `npm test`,
NativeAOT release link, Keynote run or Windows PowerPoint playback check was
added for this bounded language slice.
