# Change: independent PPTX programmable-import acceptance

## Why

Existing lossless fixtures prove selected implementation paths, but they do
not independently measure repeatability across a broad set of real edits or a
fresh Agent completing and resuming whole tasks from a packed install.

## What changes

- Define ten or more real edit intents for each of the three frozen samples.
- Run every intent three times from clean source bytes.
- Add an evaluator-owned OPC/XML/SVG/pixel oracle that does not trust runtime
  edit receipts as its verdict.
- Run three task/resume/publish prompts three times each in fresh Codex
  contexts against one packed clean install.
- Preserve all failures with their original cause.

## Non-goals

This change does not modify OfficeKit runtime, codec, wire, Help, Skills,
coverage, release metadata, or repository gates. It does not claim arbitrary
PPTX editing from success on these bounded samples.
