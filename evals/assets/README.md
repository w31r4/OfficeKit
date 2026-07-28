# PromptBench locked corpus

This directory is evaluator-side test data. It is excluded from the npm
package and is never copied into an Agent trial except for the exact declared,
read-only input files.

The initial PDF boundary set is self-authored. It contains no customer data,
production credentials, production trust anchors or private signing keys, or
third-party sample assets.
`scripts/agent-eval-corpus-fixtures.py` is the reviewed source recipe; the
checked-in `integrity.json` pins the byte hashes actually used by PromptBench.

Use the evaluator Python to verify it:

```bash
/path/to/python3 scripts/agent-eval-corpus-fixtures.py verify
```

The `pdf/signing/docmdp-p1-final.pdf` fixture is a real self-authored
certification signature with `/Perms` → `/DocMDP` permission `P=1`. Its paired
`pdf/signing/test-pki/root.pem` is a public test trust root, not a private key.
The normal verifier checks its locked bytes, visible and metadata canaries,
ByteRange/CMS presence, and DocMDP structure without requiring pyHanko.

Refreshing fixtures is an intentional corpus update: regenerate the files,
review their structure and provenance, then commit the revised assets and
integrity manifest together. Generating the DocMDP fixture additionally needs
an explicitly selected, policy-authorized managed pyHanko interpreter; private
root and signer material exists only in a temporary directory during signing:

```bash
OFFICE_KIT_PROMPTBENCH_SIGNING_PYTHON=/path/to/managed/python3 \
  /path/to/evaluator-python3 scripts/agent-eval-corpus-fixtures.py generate
```

`--signing-python /path/to/managed/python3` is equivalent. Do not treat the
fixture passwords, public test root, or any future PKI material as production
credentials.

To refresh only the time-bound signed fixture without changing the other locked
boundary assets, replace `generate` with `refresh-docmdp` in that command.
