# OfficeKit PPTX lossless-editing evidence

Status: implementation candidate, not the final completion report. This report
records the evidence at commit `894a8091a6888de42903c32f35651b43ef9e7e61`.
The persistent lossless-editing Goal remains open until the Windows PowerPoint
host lane and the remaining release conditions are complete.

## Corpus

The three external decks are immutable inputs. The fourth file is a repository
owned SmartArt canary used because none of the external decks contains SmartArt.
The authoritative inventories and source hashes are in
`manifest.v1.json`; no output is used as an input.

| sample | source bytes | source SHA-256 | slides | parts / relationships | masters / layouts / themes | charts / media / embeddings | notes / diagrams / OLE | declared targets |
| --- | ---: | --- | ---: | ---: | --- | --- | --- | ---: |
| 算秩未来 2026 | 6,114,228 | `b34ddad8cf8bbd083b60e07f8488267b1a0e4199db422468faa0eeb5d83e1762` | 21 | 595 / 660 | 2 / 34 / 4 | 4 / 50 / 25 | 17 / 0 / 84 | 7 |
| 蓝灰酸性模板 | 13,951,095 | `558ce85c0d64cd2a06faf88d6a4aa331e8cd4c685c59101c835ded2fbc87696d` | 19 | 548 / 648 | 1 / 11 / 3 | 0 / 44 / 0 | 18 / 0 / 0 | 3 |
| 麦肯锡风 customer loyalty | 47,589 | `e0bfb89454f51c400ac03797c255aa93919328ff8dba36fe414e5bcfed0536c5` | 8 | 60 / 57 | 1 / 11 / 1 | 0 / 8 / 0 | 0 / 0 / 0 | 1 |
| SmartArt canary | 13,258 | `bcb469d5b586f4fd8f562b918c8d9f04ef500cd6289728683c10ee2ced7be367` | 4 | 27 / 24 | 1 / 1 / 0 | 0 / 0 / 0 | 1 / 4 / 0 | 1 |

The machine-readable source of truth is kept in
`manifest.v1.json`, `evidence.v1.json`, `source-continuation-native.v1.json`,
`source-agent-continuation.v1.json`, and
`source-component-reuse.v1.json`.

## What is proven

- All four clean sources have byte-identical no-op exports.
- The frozen Edit Plan corpus contains 36 clean-source runs over 12 bounded
  targets, repeated three times with deterministic plans, exact non-target
  package parts, masked target XML, and successful second import.
- Each external deck has source-derived slide reuse followed by typed
  continuation. Original pages remain pixel-identical in the available
  LibreOffice/Poppler renderer, the appended page is non-blank, and the source
  remains unchanged.
- Each external deck has a passed atomic component batch using only codec-issued
  leaf IDs:

  | sample | issued leaves | changed part | reimport | source protected |
  | --- | --- | --- | --- | --- |
  | 算秩未来 2026 | `widthEmu`, `heightEmu` | `ppt/slides/slide1.xml` | yes | yes |
  | 蓝灰酸性模板 | `text`, `text` | `ppt/slides/slide12.xml` | yes | yes |
  | 麦肯锡风 customer loyalty | `leftEmu`, `topEmu` | `ppt/slides/slide1.xml` | yes | yes |

- Three independent fresh-workspace Agent runs pass the bounded native-leaf
  slice: one title edit, one subtitle edit, and one safe image-frame edit.
  Each changes one SlidePart, preserves every non-target OPC part, reimports,
  protects the source, and reports no new review issue.
- A real three-session task proves SmartArt edit → review → commit → resume /
  re-inspect → SlidePart title edit → review → commit → resume / verify →
  publish. The resumed process rebuilds the node index from reviewed bytes;
  it does not restore a JavaScript heap.

## Gate evidence

| gate | result | evidence |
| --- | --- | --- |
| Fast gate | passed, 34/34 | local `npm test` |
| Slow gate | passed, 85/85 | local `npm run test:slow` |
| OfficeKit C# | passed, 431/431 | `npm run test:office-kit-dotnet` |
| OfficeBridge | passed, 5/5 | `npm run test:office-bridge-dotnet` |
| Proto | passed | `npm run proto:check` |
| WASM reproducibility | passed, 39 audited files; 38 runtime files; 15,615,723 bytes | `npm run verify:office-kit-build` |
| Package | passed, 745 files; 36.5 MB tarball; 54.0 MB unpacked | `npm run test:pack` |
| API / offline release metadata | passed | `docs:api`; `release:check --skip-network --skip-commands` in hosted lane |
| Hosted slow CI | passed | [run 32542799424](https://github.com/w31r4/OfficeKit/actions/runs/32542799424), SHA `894a8091` |

The local full `release:check` reached every code, package, .NET, and metadata
check but could not run `npm whoami`: this machine has no npm publishing
credential. That is a publication prerequisite, not a codec result.

## Boundaries that remain open

- Real Microsoft PowerPoint desktop acceptance has not been performed. The
  Windows lane must still open, browse, save a copy, and reconnect the edited
  decks without a repair prompt. macOS rendering and mocks are not substitutes.
- The native-leaf and component-batch APIs are bounded. Opaque or topology-
  bearing groups, arbitrary Master/Layout/theme rewrites, animation, complex
  SmartArt authoring, OLE internals, and unsupported relationship graphs remain
  read-only or fail closed.
- Visual review on the current macOS host is unavailable for the black-box
  Agent runs; structural/package/render-oracle evidence is not an aesthetic
  judgement.
- A public npm publication still needs configured npm authentication and the
  release owner's tag/publication action.

## Next acceptance step

Run the self-hosted Windows Office lane against this SHA with the three external
decks and record manifest upload, two-deck isolation, unsaved edit, selection,
single-page image review, explicit save, reconnect/disconnect, source protection,
and unsupported-capability refusal. Only after that evidence is attached should
the OpenSpec 6.2/6.5/6.6 items be closed and the persistent Goal be marked
complete.
