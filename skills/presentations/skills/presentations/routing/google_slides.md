# Google Slides routing

Create and verify a local `.pptx` with the Presentations Skill first. OfficeKit
does not upload files, create cloud decks, or operate a Google Drive. After
the local deck passes semantic and render QA, the user or another host may
import it into Google Slides.

For an existing native Google Slides deck, obtain a local export or a user
provided reference before editing. Preserve that input and return a distinct
local output unless the user explicitly asks for an in-place host operation.

Return the verified `.pptx` path, SHA-256, slide locators, and evidence
envelope. If a cloud link is needed, state that import is a separate host
step.
