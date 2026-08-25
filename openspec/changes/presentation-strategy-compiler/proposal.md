## Why

OfficeKit already has durable authoring plans, free Compose primitives, native
motion, review, and resumable tasks, but its default Presentation route treats
design mostly as an abstract grammar. It does not make the communication job,
delivery context, scenario craft, and chosen visual direction mandatory inputs,
so a capable runtime can still produce repetitive or weakly composed decks.

## What Changes

- Publish an OfficeKit presentation doctrine that defines a deck as a
  communication activity, editable document, playback experience, and native
  artifact, with factual, communication, narrative, cognitive, visual, and
  native/run-time quality layers.
- Extend authoring-plan v1 with additive communication-job, medium-fit,
  after-use, scenario, and chosen-direction fields while preserving existing
  tasks and the plan schema identifier.
- Add seven clean-room scenario guides and require one primary scenario, an
  optional secondary scenario, one of four design sources, and a deck-specific
  direction before composition.
- Keep scenarios orthogonal to the existing design mechanisms and preserve
  user templates, brand rules, and references as higher design authorities.
- Make Presentation review report bounded communication, narrative, cognitive,
  and visual-risk evidence without pretending to verify facts or aesthetic
  quality automatically.
- Replace examples that teach a universal rounded-card treatment with distinct
  analysis, management, and brand visual grammars, then exercise them through
  three focused real decks rather than a new benchmark matrix.

## Capabilities

### New Capabilities

- `presentation-strategy-authoring`: Communication-first authoring-plan fields,
  medium-fit behavior, chosen visual direction, task/resume descriptors, and
  the public presentation doctrine.
- `presentation-scenario-design`: Clean-room seven-scenario policy, four design
  sources, scenario/mechanism composition, progressive Skill routing, and
  deck-specific design-grammar requirements.
- `presentation-strategy-review`: Plan-bound communication, narrative,
  cognitive, and visual-risk review evidence with honest fact and aesthetic
  boundaries.

### Modified Capabilities

None. These behaviors have not been archived into canonical repository specs.

## Impact

- Additive authoring-plan v1 validation, task descriptors, REPL resume state,
  Presentation review data, Help metadata, and generated API documentation.
- Presentations Skill routing, shared style guidance, scenario references,
  examples, README positioning, coverage, and release evidence.
- Target release is `0.9.0`; Office wire protocol remains 2 and no C# Codec,
  PDF, Spreadsheet, Document, Live Add-in, provider, or template asset format
  changes are included.
- Kimi-derived files remain research-only and are neither copied nor packaged.
