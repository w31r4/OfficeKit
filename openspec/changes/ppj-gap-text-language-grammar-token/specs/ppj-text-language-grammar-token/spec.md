# PPJ text language grammar token

## ADDED Requirements

### Requirement: resolve a typed language token

The authored compiler MUST accept a `grammarTokenRef` for text language when
the referenced token has kind `string`, and MUST validate the resolved value
as a bounded BCP-47 language tag before lowering it to `a:rPr/@lang`.

#### Scenario: formal theme language uses a token

- **WHEN** a formal `text.language` precedence rule selects `theme` and the
  theme text style contains `{ "token": "language" }`
- **AND** the grammar declares `language` with kind `string` and value
  `ar-SA`
- **THEN** authored output writes `ar-SA` to the run language owner
- **AND** a second projection recovers `ar-SA`

### Requirement: reject invalid token values

The compiler MUST reject a language token with a non-string kind or a string
that is not a valid bounded BCP-47 language tag.

#### Scenario: wrong language token fails closed

- **WHEN** `text.language` references a `number` token or a malformed tag
- **THEN** authored compilation fails with a grammar-token or presentation
  language diagnostic
