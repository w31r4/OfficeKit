# PPJ text language

## Why

PPJ declares BCP-47 run language but the authored compiler rejects it and the
native wire silently defaults every new run to `en-US`. This makes a documented
field non-functional and weakens Chinese/Latin font selection, host shaping,
spell checking and accessibility.

## What Changes

- Carry bounded run and default-run language through additive wire-v2 fields.
- Read and write direct DrawingML `lang` attributes.
- Compile and project PPJ `textStyle.language`.
- Issue an exact source-bound `fontLanguage` leaf for canonical direct run
  language attributes.
- Keep inherited, malformed and extension-bearing language state source owned.

## Impact

- Additive protobuf fields; protocol version remains 2.
- The PPJ schema is unchanged because `language` already exists.
- One existing integrated PPJ test covers authored, projected and capability-
  bound continuation behavior.
