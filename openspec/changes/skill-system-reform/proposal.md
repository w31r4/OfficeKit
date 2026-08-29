## Why

OfficeKit has accumulated capable Presentation primitives, but their guidance is
split across a long router, API prose, examples, scenario notes, and template
instructions. A new Agent can find a method without understanding when to use
it, while future primitive changes have no explicit list of Skills, references,
Help, examples, and gates that must move together. Kimi's presentation package
shows a cleaner separation: one short route, task-specific references, a
machine-readable format contract, and independent typography/shape/scenario
guidance.

## What Changes

- **BREAKING (Skill surface only)** Reorganize the OfficeKit Presentation
  guidance around a short route and a single production spine; remove repeated
  workflow, API, and delivery prose from secondary files rather than adding
  another parallel guide.
- Publish a clean-room map of Kimi-observed surfaces, explicitly separating
  portable presentation principles, technical format/API facts, provider or
  product-specific behavior, and facts that OfficeKit must not copy.
- Add a concise OfficeKit presentation primitive reference analogous to Kimi's
  `pptd.md`/`shapes.md`: semantic families, capability boundaries, source-bound
  behavior, and links to the authoritative API/Help entries. It is a navigation
  index, not a second API specification.
- Add a typography reference analogous to Kimi's `fonts.md`, using installed
  font evidence and role-based selection rules without inventing universal
  palettes or copying Kimi font prose.
- Add `skill-update`, a host-neutral Skill and deterministic checker that maps
  changed runtime/proto/Help/API/Skill/example paths to the required update
  surfaces and reports stale or missing links. It does not edit files or load
  heavy runtimes automatically.
- Add a versioned primitive-impact manifest as the single maintenance contract
  for future Presentation primitive changes. It records source owners,
  consumer Skills, Help/API generation, examples, focused tests, and release
  evidence.
- Keep Kimi's local files research-only and clean-room: no Kimi source text,
  private endpoints, `.pptd` runtime, remote templates, or proprietary assets
  enter the OfficeKit package.

## Capabilities

### New Capabilities

- `skill-update-workflow`: A reusable Skill and checker for propagating a
  primitive or workflow change through its documented consumer surfaces.
- `presentation-primitive-surface`: A concise semantic index for OfficeKit
  Presentation primitives, their authority boundaries, and Agent discovery
  paths.
- `presentation-typography-guidance`: Role-based font selection and fallback
  guidance for native, imported, and rendered presentations.
- `presentation-skill-routing`: The short Presentation router and progressive
  reference-loading contract.
- `presentation-template-creator-integration`: Shared primitive/typography
  guidance and capability evidence used by the Presentation Template Creator.

### Modified Capabilities

None. The existing authoring and template behavior is represented by the new
progressive-routing and Creator-integration contracts without changing the
OfficeKit JavaScript or wire APIs.

## Impact

- `skills/presentations/skills/presentations/` route, references, task files,
  examples, and manifest metadata.
- `skills/presentation-template-creator/` guidance and packaging metadata.
- New `skills/skill-update/` plugin, checker script, manifest, and marketplace
  entry; `officekit init` includes it as a lightweight maintenance Skill.
- `src/help/index.mjs`, generated `docs/api.md`, and a small impact manifest;
  no Office wire, codec, or public JavaScript API change.
- Narrow Skill/reference synchronization checks and one existing presentation
  smoke; no benchmark matrix or broad regression suite.
