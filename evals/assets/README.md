# PromptBench locked corpus

This directory is evaluator-side test data. It is excluded from the npm
package and is never copied into an Agent trial except for the exact declared,
read-only input files.

The initial PDF boundary set is self-authored. It contains no customer data,
production credentials, trusted signing keys, or third-party sample assets.
`scripts/agent-eval-corpus-fixtures.py` is the reviewed source recipe; the
checked-in `integrity.json` pins the byte hashes actually used by PromptBench.

Use the bundled evaluator Python to refresh or verify it:

```bash
/path/to/python3 scripts/agent-eval-corpus-fixtures.py verify
```

Refreshing fixtures is an intentional corpus update: regenerate the files,
review their structure and provenance, then commit the revised assets and
integrity manifest together. Do not treat the fixture passwords or any future
PKI material as production credentials.
