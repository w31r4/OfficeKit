# OfficeKit repository guide

OfficeKit is a local, agent-facing Office and PDF toolkit. It contains the
JavaScript object model and CLI, the OfficeKit Codec C# codec and WASM runtime,
the Excel Live Add-in, and the portable Skill packages that teach an Agent how
to use those capabilities.

## Source-of-truth boundaries

- `src/` is the public JavaScript runtime and CLI. Keep imports leaf-oriented;
  the root entry must not eagerly initialize Office WASM, MuPDF, providers, or
  the Excel bridge.
- `native/OfficeKit/` is the C# codec source. `runtime/office-kit/` is audited
  generated WASM output; rebuild it with the checked-in build commands instead
  of editing generated files by hand.
- `proto/` is the versioned wire contract. Run `npm run proto:check` after a
  protocol change and do not change the Office wire version for a Skill-only
  change.
- `skills/<plugin>/` is the canonical plugin source. `.codex-plugin` manifests
  describe the Codex package surface; the root `.claude-plugin/marketplace.json`
  is a thin Claude marketplace index over the same Skill trees.
- `apps/excel-addin/` and `apps/powerpoint-addin/` are local Live host
  adapters. They are not Office file codecs and must remain lazy from root
  imports. Excel keeps its compatibility commands; PowerPoint uses the common
  Live bridge and typed `officekit live` operations. Word Live is only a future
  adapter contract.
- `openspec/` contains change proposals, specs, designs, and task checklists.
- `test/` contains gates and artifact fixtures. `tmp/` is disposable QA output
  and must never become a package or source reference.

## Working rules

1. Preserve source files and user-provided references. Never overwrite an input
   artifact or silently switch to a different authoring engine.
2. Unsupported imported topology must remain opaque or fail closed. Do not
   flatten it merely to make an edit succeed.
3. Keep Skills host-neutral: no Codex-only tools, thread identifiers, MCP
   names, or required image-generation tools. Describe capabilities and
   evidence, not a particular host's message syntax.
4. Keep provider packs explicit and lazy. No lifecycle download, network fetch,
   or large specialist runtime belongs in a root import or ordinary smoke test.
5. Add a test, update the relevant Skill/docs/coverage entry, and run the
   narrowest affected gates before a full `npm test`.
6. Generated API docs, protobuf bindings, WASM, manifests, SBOMs, and package
   inventories are release evidence. Regenerate them with their scripts and
   review the resulting diff.

## Useful checks

```sh
npm test
npm run docs:api
npm run proto:check
npm run build:office-kit
npm run verify:office-kit-build
node test/reference-skills.mjs
node test/claude-plugin.mjs
```

For a package-only change, also run `npm pack --dry-run --json` and
`node test/package-contents.mjs`. For a Skill change, run the portability and
reference-sync gates. For a PDF provider change, run the provider-specific
contract tests and document any environment-dependent skip.
