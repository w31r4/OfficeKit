## Context

Presentation authoring already supports PNG, JPEG, GIF, and safe SVG placement, source-bound image replacement, rendering, and accessibility metadata. What is missing is the acquisition boundary before placement: lawful discovery, safe retrieval, task-local provenance, resumability, and proof that the delivered PPTX uses the registered bytes.

The root package must stay lazy. Ordinary imports, `officekit init`, template search, and non-image tasks cannot load provider data or perform network work. Task schema v2 is already durable and private; image assets can live beneath a task without becoming a new artifact kind or manifest field.

## Goals / Non-Goals

**Goals:**

- Give an Agent a deterministic CLI for image candidates, acquisition, listing, and media audit.
- Preserve image bytes and provenance inside the active task so a new context can resume without repeating a search.
- Enforce a narrow license policy and secure HTTPS download boundary before any remote bytes reach Presentation code.
- Let existing image placement APIs consume a `FileBlob` without large hand-built data URLs.
- Teach the Presentations Skill when imagery is useful, how to select it, and how to review crop, clarity, accessibility, attribution, and delivery evidence.

**Non-Goals:**

- No independent image Skill, MCP server, vector database, image-generation provider, legal guarantee, or automatic visual taste model.
- No Unsplash, Pexels, Pixabay, WebP, remote SVG, arbitrary URL guessing, transcoding, hotlinking, or silent provider fallback.
- No task schema migration, new public JavaScript subpath, Office wire change, or C# codec work.
- No benchmark matrix or provider conformance suite.

## Decisions

### Keep image acquisition behind a leaf-loaded CLI

`src/cli/officekit.mjs` dynamically imports `src/images/cli.mjs` only for the `image` command. Provider modules and the Lucide icon bundle are imported inside the relevant search request. This preserves the root-import and initialization boundary.

Alternative considered: a separate Skill or MCP server. Rejected because image discovery is a deterministic package primitive needed by the Presentation workflow, not a second Agent experience or long-running service.

### Use provider adapters, not the webfetch downloader or federation defaults

OfficeKit pins `webfetch-core@0.1.5` and calls only its exported Openverse/Wikimedia providers and license normalization utilities. It passes an explicit provider list and never invokes package defaults, cache, downloader, browser fallback, MCP, or CLI. Lucide is loaded from `@iconify-json/lucide@1.2.126` and rendered deterministically from Iconify path data.

Alternative considered: writing both provider integrations from scratch. Rejected because webfetch-core already normalizes their candidate and license shapes. Its download behavior is deliberately not reused because OfficeKit needs stricter DNS, redirect, type, byte, and pixel controls.

### Make candidates opaque and task-bound

Each search writes one private evidence record under `evidence/images/searches/<search-id>/search.json`. Public results contain a `candidateRef`, preview/source metadata, dimensions, license facts, and `selectionMade: false`, but omit the acquisition URL. `image add --candidate` resolves the ref only inside the named task and reads the stored provider candidate.

This makes search resumable and prevents callers from substituting a new URL behind an approved candidate. A provider error remains visible in the search report; results from another requested provider are not described as a fallback.

### Store immutable content-addressed assets outside the task manifest

Acquired bytes are stored as `assets/images/<sha256>.<ext>` with a sibling receipt. Re-adding identical bytes is idempotent. Assets are read-only, receipts and search/audit records are private, and writes use task-contained temporary files followed by atomic rename. `image list` scans receipts, so task schema v2 and artifact kinds remain unchanged.

### Own the remote download security boundary

The downloader accepts HTTPS only, rejects credentials and unsafe hostnames, resolves DNS before connecting, rejects any private/link-local/loopback/metadata address, and pins the selected public address into the HTTPS connection. Every redirect is revalidated and at most three are followed. Streaming stops above 20 MiB. After download, magic bytes, declared MIME, actual dimensions, 40 MP, and 16,384-pixel edge limits are checked before publishing the asset.

Remote acquisition accepts PNG, JPEG, and GIF only. Local SVG and Lucide SVG continue through existing Presentation safe-SVG validation on export; remote SVG is refused.

### Treat licenses as provenance evidence, not legal certification

The allowlist is Public Domain, CC0, CC BY, Lucide ISC, user-provided/generated/permission declarations, and official press kits. CC BY requires author, license URL, and a generated credit line. ShareAlike, NonCommercial, NoDerivatives, and unknown status are rejected. Openverse records are labeled `provider-declared`; Wikimedia records preserve machine-readable source metadata. Audit sidecars never replace visible attribution duties.

### Convert FileBlob at the Presentation boundary

`ImageElement` accepts either the existing `dataUrl` or a `FileBlob`. It validates that only one byte source was supplied, checks the image MIME, and converts the blob bytes into the existing canonical data URL representation immediately. Compose passes the same property through. The serialized model and codec remain unchanged.

### Audit the bytes that actually shipped

`image audit` opens the PPTX package, hashes every `ppt/media/*` part, and joins hashes to task receipts. It reports registered assets used by the deck, registered but unused assets, unregistered media, and visible attribution obligations. An optional deterministic `.sources.json` sidecar contains the same evidence.

## Risks / Trade-offs

- **Provider metadata can be incorrect** → preserve provider/source evidence, label confidence, reject unknown licenses, and avoid claiming legal verification.
- **Search APIs can drift or throttle** → isolate adapters, return provider reports, keep ordinary tests offline with injected providers, and never switch sources silently.
- **DNS pinning differs across Node/platform networking** → use the built-in HTTPS client with an explicit pinned lookup and test the boundary independently from provider search.
- **Iconify metadata does not encode semantic aliases for every intent** → use deterministic name/token matching and return no candidates successfully when confidence is poor.
- **A PPTX may contain pre-existing unregistered media** → audit reports it without deleting or rewriting the deck.
- **Embedding bytes increases task and package sizes** → content-address duplicate assets, cap individual images, and keep the icon collection compressed inside one dependency.

## Migration Plan

1. Land the OpenSpec and additive runtime/CLI implementation on an isolated branch.
2. Add exact dependencies, notices, package inventory, and `1.1.0` release metadata.
3. Update the Presentations Skill and reference-sync copies without changing other Skills.
4. Run narrow offline contracts and one real single-slide dogfood, then the final package/release gates once.
5. Rollback is removal of the lazy command and dependencies; existing PPTX, tasks, and image APIs remain compatible because no persisted schema or wire contract changes.

## Open Questions

None for v1. Additional keyed providers, remote SVG, image generation, and broader document consumers require separate changes.
