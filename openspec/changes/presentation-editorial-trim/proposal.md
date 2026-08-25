## Why

OfficeKit already locks facts and compresses presentation copy, but its current
guidance is too general to catch the sentence patterns that make otherwise good
decks sound machine-written. The existing `Trim CN Tech Doc` Skill has a strong
plain-language rule set, while its section-by-section approval flow and
document-oriented deletion policy do not fit slide titles, live delivery,
speaker notes, or one-shot deck creation.

## What Changes

- Add a small, host-neutral `presentation-editorial-trim` Skill inside the
  Presentations plugin. The main Presentations workflow invokes it during
  authoring; users may also invoke it explicitly to polish an existing deck.
- Adapt the proven trim patterns for slide titles, visible body copy, labels,
  sources, and speaker notes. Preserve facts, citations, qualifiers, proper
  nouns, user wording, and page scope.
- Use two bounded passes: editorial shaping before composition and page-fit
  editing after render. The second pass may shorten copy or change hierarchy,
  but must not silently change evidence or shrink text below the plan floor.
- Make delivery mode part of the editorial decision: live decks move support
  into notes, reader decks retain necessary qualifiers, and hybrid decks keep
  the visible argument self-contained.
- Treat false contrast, defensive negation, empty signposts, abstract noun
  chains, repeated three-part phrasing, slogan fragments, and repeated title
  forms as review targets rather than mechanical string bans.
- Reuse the existing authoring plan `editorial` object and task/review flow. No
  codec, wire, public JavaScript API, or new document format is introduced.
- Keep the change independent from the external `Trim CN Tech Doc` installation;
  OfficeKit ships the smaller presentation-specific rules it needs.

## Capabilities

### New Capabilities

- `presentation-editorial-trim`: Presentation-specific copy shaping, AI-pattern
  removal, evidence preservation, delivery-aware compression, and deck-wide
  voice review.

### Modified Capabilities

None.

## Impact

- Adds one focused Skill under `skills/presentations/skills/` and a compact
  reference/evaluation corpus within the Presentations plugin.
- Updates the create, create-from-template, edit-existing, continue, and review
  routes to call the editorial pass at the correct scope.
- Updates Presentations package inventory, Skill smoke checks, and coverage
  documentation. Office wire protocol and runtime exports remain unchanged.
