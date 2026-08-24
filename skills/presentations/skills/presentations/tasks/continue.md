# Continue a durable presentation task

Use this route when the user wants to resume prior OfficeKit work.

## 1. Find the task

Run `officekit tasks --json` in the intended workspace. Match goal, inputs,
artifacts, source hashes, and next action. Ask when more than one task matches.

Open the selected task with `officekit repl <task-id> --file <cell.mjs>`.

## 2. Recover durable state

Read:

- `session.ready.task.plan` for compact plan identity and state;
- `await ctx.plan()` for the full validated plan;
- `ctx.task.commit` and restored artifact revisions;
- pending review failures and `next`.

The new process has no prior JavaScript heap. Reimport the latest reviewed PPTX
and rebuild all node/capability indexes with `inspect()`.

## 3. Continue from the reviewed revision

Follow the plan's `recipe` and `nextAction`. If intent or design changes, write
an updated plan with its current SHA-256. The task becomes `working` until a new
artifact commit binds that plan.

For local edits, preserve unchanged page copy and design roles and pass exact
`changedPageIds` to review. Commit each meaningful reviewed phase. Publish only
the current reviewed commit.

Read the installed OfficeKit Skill's `references/repl.md` for crash recovery,
immutable inputs, and publication rules.
